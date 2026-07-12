using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.CSISapModel.FrameObject;
using ExcelCSIToolBox.Data.CSISapModel.Truss;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel.Truss
{
    public sealed class CsiHoweTrussGenerationService : ICsiTrussGenerationService
    {
        private const int DefaultBayCount = 6;
        private const int MaxBayCount = 100;
        private const double DefaultSpan = 12000.0;
        private const double DefaultHeight = 3000.0;
        private const string SlopeModeNone = "None";
        private const string SlopeModeMono = "Mono";
        private const string SlopeModeGable = "Gable";
        private const string MonoSlopeLeftToRight = "LeftToRight";
        private const string MonoSlopeRightToLeft = "RightToLeft";
        private const string LoadTargetTopChord = "TopChord";
        private const string LoadTargetBottomChord = "BottomChord";
        private const string LoadTargetChord = "Chord";
        private const string LoadTargetWeb = "Web";
        private const string LoadTargetAll = "All";
        private const string TrussTypeHowe = "Howe";
        private const string TrussTypePratt = "Pratt";

        public OperationResult<HoweTrussResultDto> Generate(
            ICSISapModelConnectionService service,
            HoweTrussRequestDto request)
        {
            if (service == null)
            {
                return OperationResult<HoweTrussResultDto>.Failure("active CSI model is not available.");
            }

            HoweTrussRequestDto input = Normalize(request);
            double baySpacing = input.Span / input.BayCount;
            var chordNames = new List<string>();
            var topChordNames = new List<string>();
            var bottomChordNames = new List<string>();
            var webNames = new List<string>();
            var frames = new List<FrameAddRequestDto>();

            AddChordFrames(input, baySpacing, frames, chordNames, topChordNames, bottomChordNames);
            AddVerticalFrames(input, baySpacing, frames, webNames);
            AddDiagonalFrames(input, baySpacing, frames, webNames);

            OperationResult<FrameAddBatchResultDto> addResult = service.AddFrameObjects(new FrameAddBatchRequestDto
            {
                Frames = frames
            });

            var result = new HoweTrussResultDto
            {
                BayCount = input.BayCount,
                Span = input.Span,
                BaySpacing = baySpacing,
                TrussType = input.TrussType,
                SlopeMode = input.SlopeMode,
                Slope = input.Slope,
                MonoSlopeDirection = input.MonoSlopeDirection,
                DistributedLoadPattern = input.DistributedLoadPattern,
                DistributedLoadDirection = input.DistributedLoadDirection,
                DistributedLoadValue1 = input.DistributedLoadValue1,
                DistributedLoadValue2 = input.DistributedLoadValue2,
                DistributedLoadTarget = input.DistributedLoadTarget,
                ChordFrameNames = new List<string>(),
                WebFrameNames = new List<string>(),
                FailureReasons = new List<string>()
            };

            if (!addResult.IsSuccess || addResult.Data == null)
            {
                result.FailureReasons.Add(addResult.Message ?? $"Failed to add {input.TrussType} truss frame objects.");
                return OperationResult<HoweTrussResultDto>.Success(result);
            }

            result.AddedFrameCount = addResult.Data.SuccessCount;
            foreach (FrameAddResultDto item in addResult.Data.Results)
            {
                if (item == null || !item.Success || string.IsNullOrWhiteSpace(item.FrameName))
                {
                    if (item != null && !string.IsNullOrWhiteSpace(item.FailureReason))
                    {
                        result.FailureReasons.Add(item.FailureReason);
                    }
                    continue;
                }

                if (ContainsName(chordNames, item.FrameName))
                {
                    result.ChordFrameNames.Add(item.FrameName);
                }
                else
                {
                    result.WebFrameNames.Add(item.FrameName);
                }
            }

            AssignDistributedLoadIfRequested(service, input, result, topChordNames, bottomChordNames, chordNames, webNames);

            if (result.WebFrameNames.Count > 0)
            {
                OperationResult releaseResult = service.SetFrameReleases(
                    result.WebFrameNames,
                    new[] { false, false, false, false, true, true },
                    new[] { false, false, false, false, true, true });

                if (releaseResult.IsSuccess)
                {
                    result.ReleasedWebMemberCount = result.WebFrameNames.Count;
                }
                else
                {
                    result.FailureReasons.Add(releaseResult.Message);
                }
            }

            result.Success = result.FailureReasons.Count == 0;
            return OperationResult<HoweTrussResultDto>.Success(result);
        }

        private static void AddChordFrames(
            HoweTrussRequestDto input,
            double baySpacing,
            List<FrameAddRequestDto> frames,
            List<string> chordNames,
            List<string> topChordNames,
            List<string> bottomChordNames)
        {
            for (int i = 0; i < input.BayCount; i++)
            {
                string bottomName = $"{input.NamePrefix}_BC{i + 1:000}";
                string topName = $"{input.NamePrefix}_TC{i + 1:000}";
                chordNames.Add(bottomName);
                chordNames.Add(topName);
                bottomChordNames.Add(bottomName);
                topChordNames.Add(topName);

                double xi = input.StartX + i * baySpacing;
                double xj = input.StartX + (i + 1) * baySpacing;
                frames.Add(CreateFrame(bottomName, input.ChordPropName, xi, input.StartY, input.StartZ, xj, input.StartY, input.StartZ));
                frames.Add(CreateFrame(topName, input.ChordPropName, xi, input.StartY, TopZAt(input, xi), xj, input.StartY, TopZAt(input, xj)));
            }
        }

        private static void AssignDistributedLoadIfRequested(
            ICSISapModelConnectionService service,
            HoweTrussRequestDto input,
            HoweTrussResultDto result,
            List<string> topChordNames,
            List<string> bottomChordNames,
            List<string> chordNames,
            List<string> webNames)
        {
            if (string.IsNullOrWhiteSpace(input.DistributedLoadPattern) ||
                input.DistributedLoadValue1 == 0 && input.DistributedLoadValue2 == 0)
            {
                return;
            }

            List<string> targetNames = GetDistributedLoadTargets(input.DistributedLoadTarget, topChordNames, bottomChordNames, chordNames, webNames);
            if (targetNames.Count == 0)
            {
                result.FailureReasons.Add($"No {input.TrussType} truss frame objects matched the requested distributed load target.");
                return;
            }

            OperationResult loadResult = service.AssignFrameDistributedLoad(
                targetNames,
                input.DistributedLoadPattern,
                input.DistributedLoadDirection,
                input.DistributedLoadValue1,
                input.DistributedLoadValue2);

            if (loadResult.IsSuccess)
            {
                result.LoadedFrameCount = targetNames.Count;
            }
            else
            {
                result.FailureReasons.Add(loadResult.Message);
            }
        }

        private static List<string> GetDistributedLoadTargets(
            string target,
            List<string> topChordNames,
            List<string> bottomChordNames,
            List<string> chordNames,
            List<string> webNames)
        {
            if (string.Equals(target, LoadTargetBottomChord, StringComparison.OrdinalIgnoreCase))
            {
                return new List<string>(bottomChordNames);
            }

            if (string.Equals(target, LoadTargetChord, StringComparison.OrdinalIgnoreCase))
            {
                return new List<string>(chordNames);
            }

            if (string.Equals(target, LoadTargetWeb, StringComparison.OrdinalIgnoreCase))
            {
                return new List<string>(webNames);
            }

            if (string.Equals(target, LoadTargetAll, StringComparison.OrdinalIgnoreCase))
            {
                var allNames = new List<string>(chordNames);
                allNames.AddRange(webNames);
                return allNames;
            }

            return new List<string>(topChordNames);
        }

        private static void AddVerticalFrames(
            HoweTrussRequestDto input,
            double baySpacing,
            List<FrameAddRequestDto> frames,
            List<string> webNames)
        {
            for (int i = 0; i <= input.BayCount; i++)
            {
                string name = $"{input.NamePrefix}_V{i + 1:000}";
                webNames.Add(name);
                double x = input.StartX + i * baySpacing;
                frames.Add(CreateFrame(name, input.WebPropName, x, input.StartY, input.StartZ, x, input.StartY, TopZAt(input, x)));
            }
        }

        private static void AddDiagonalFrames(
            HoweTrussRequestDto input,
            double baySpacing,
            List<FrameAddRequestDto> frames,
            List<string> webNames)
        {
            double center = input.BayCount / 2.0;
            for (int i = 0; i < input.BayCount; i++)
            {
                string name = $"{input.NamePrefix}_D{i + 1:000}";
                webNames.Add(name);

                bool leftHalf = i < center;
                bool diagonalStartsAtTop = string.Equals(input.TrussType, TrussTypePratt, StringComparison.OrdinalIgnoreCase)
                    ? !leftHalf
                    : leftHalf;

                if (diagonalStartsAtTop)
                {
                    double xi = input.StartX + i * baySpacing;
                    double xj = input.StartX + (i + 1) * baySpacing;
                    frames.Add(CreateFrame(
                        name,
                        input.WebPropName,
                        xi,
                        input.StartY,
                        TopZAt(input, xi),
                        xj,
                        input.StartY,
                        input.StartZ));
                }
                else
                {
                    double xi = input.StartX + i * baySpacing;
                    double xj = input.StartX + (i + 1) * baySpacing;
                    frames.Add(CreateFrame(
                        name,
                        input.WebPropName,
                        xi,
                        input.StartY,
                        input.StartZ,
                        xj,
                        input.StartY,
                        TopZAt(input, xj)));
                }
            }
        }

        private static double TopZAt(HoweTrussRequestDto input, double x)
        {
            double station = Math.Max(0.0, Math.Min(input.Span, x - input.StartX));
            double baseTopZ = input.StartZ + input.Height;
            if (input.Slope <= 0 || string.Equals(input.SlopeMode, SlopeModeNone, StringComparison.OrdinalIgnoreCase))
            {
                return baseTopZ;
            }

            if (string.Equals(input.SlopeMode, SlopeModeGable, StringComparison.OrdinalIgnoreCase))
            {
                double distanceFromNearestEnd = Math.Min(station, input.Span - station);
                return baseTopZ + input.Slope * distanceFromNearestEnd;
            }

            if (string.Equals(input.SlopeMode, SlopeModeMono, StringComparison.OrdinalIgnoreCase))
            {
                double run = string.Equals(input.MonoSlopeDirection, MonoSlopeRightToLeft, StringComparison.OrdinalIgnoreCase)
                    ? input.Span - station
                    : station;
                return baseTopZ + input.Slope * run;
            }

            return baseTopZ;
        }

        private static FrameAddRequestDto CreateFrame(
            string name,
            string propName,
            double xi,
            double yi,
            double zi,
            double xj,
            double yj,
            double zj)
        {
            return new FrameAddRequestDto
            {
                UserName = name,
                PropName = propName,
                Xi = xi,
                Yi = yi,
                Zi = zi,
                Xj = xj,
                Yj = yj,
                Zj = zj
            };
        }

        private static HoweTrussRequestDto Normalize(HoweTrussRequestDto request)
        {
            HoweTrussRequestDto source = request ?? new HoweTrussRequestDto();
            int bays = source.BayCount <= 0 ? DefaultBayCount : Math.Min(source.BayCount, MaxBayCount);
            double span = source.Span <= 0 ? DefaultSpan : source.Span;
            double height = source.Height <= 0 ? DefaultHeight : source.Height;

            return new HoweTrussRequestDto
            {
                BayCount = Math.Max(2, bays),
                Span = span,
                Height = height,
                StartX = source.StartX,
                StartY = source.StartY,
                StartZ = source.StartZ,
                TrussType = NormalizeTrussType(source.TrussType),
                NamePrefix = string.IsNullOrWhiteSpace(source.NamePrefix) ? DefaultNamePrefix(source.TrussType) : source.NamePrefix.Trim(),
                ChordPropName = string.IsNullOrWhiteSpace(source.ChordPropName) ? "Default" : source.ChordPropName.Trim(),
                WebPropName = string.IsNullOrWhiteSpace(source.WebPropName) ? "Default" : source.WebPropName.Trim(),
                SlopeMode = NormalizeSlopeMode(source.SlopeMode, source.Slope),
                Slope = Math.Max(0.0, source.Slope),
                MonoSlopeDirection = NormalizeMonoSlopeDirection(source.MonoSlopeDirection),
                DistributedLoadPattern = string.IsNullOrWhiteSpace(source.DistributedLoadPattern) ? null : source.DistributedLoadPattern.Trim(),
                DistributedLoadDirection = source.DistributedLoadDirection <= 0 ? 6 : source.DistributedLoadDirection,
                DistributedLoadValue1 = source.DistributedLoadValue1,
                DistributedLoadValue2 = source.DistributedLoadValue2,
                DistributedLoadTarget = NormalizeDistributedLoadTarget(source.DistributedLoadTarget)
            };
        }

        private static string NormalizeTrussType(string trussType)
        {
            return string.Equals(trussType, TrussTypePratt, StringComparison.OrdinalIgnoreCase)
                ? TrussTypePratt
                : TrussTypeHowe;
        }

        private static string DefaultNamePrefix(string trussType)
        {
            return string.Equals(trussType, TrussTypePratt, StringComparison.OrdinalIgnoreCase) ? "PT" : "HT";
        }

        private static string NormalizeSlopeMode(string slopeMode, double slope)
        {
            if (slope <= 0)
            {
                return SlopeModeNone;
            }

            if (string.Equals(slopeMode, SlopeModeMono, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(slopeMode, "Monoslope", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(slopeMode, "OneSide", StringComparison.OrdinalIgnoreCase))
            {
                return SlopeModeMono;
            }

            if (string.Equals(slopeMode, SlopeModeGable, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(slopeMode, "Double", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(slopeMode, "Middle", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(slopeMode, "TwoSide", StringComparison.OrdinalIgnoreCase))
            {
                return SlopeModeGable;
            }

            return SlopeModeGable;
        }

        private static string NormalizeMonoSlopeDirection(string direction)
        {
            if (string.Equals(direction, MonoSlopeRightToLeft, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(direction, "Right", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(direction, "FromRight", StringComparison.OrdinalIgnoreCase))
            {
                return MonoSlopeRightToLeft;
            }

            return MonoSlopeLeftToRight;
        }

        private static string NormalizeDistributedLoadTarget(string target)
        {
            if (string.Equals(target, LoadTargetBottomChord, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(target, "Bottom", StringComparison.OrdinalIgnoreCase))
            {
                return LoadTargetBottomChord;
            }

            if (string.Equals(target, LoadTargetChord, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(target, "Chords", StringComparison.OrdinalIgnoreCase))
            {
                return LoadTargetChord;
            }

            if (string.Equals(target, LoadTargetWeb, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(target, "Webs", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(target, "Brace", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(target, "Braces", StringComparison.OrdinalIgnoreCase))
            {
                return LoadTargetWeb;
            }

            if (string.Equals(target, LoadTargetAll, StringComparison.OrdinalIgnoreCase))
            {
                return LoadTargetAll;
            }

            return LoadTargetTopChord;
        }

        private static bool ContainsName(List<string> names, string value)
        {
            foreach (string name in names)
            {
                if (string.Equals(name, value, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }
    }
}
