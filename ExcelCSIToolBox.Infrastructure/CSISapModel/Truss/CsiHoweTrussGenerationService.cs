using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.CSISapModel.FrameObject;
using ExcelCSIToolBox.Data.CSISapModel.Truss;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel.Truss
{
    public sealed class CsiHoweTrussGenerationService
    {
        private const int DefaultBayCount = 6;
        private const int MaxBayCount = 100;
        private const double DefaultSpan = 12000.0;
        private const double DefaultHeight = 3000.0;

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
            var webNames = new List<string>();
            var frames = new List<FrameAddRequestDto>();

            AddChordFrames(input, baySpacing, frames, chordNames);
            AddVerticalFrames(input, baySpacing, frames, webNames);
            AddHoweDiagonalFrames(input, baySpacing, frames, webNames);

            OperationResult<FrameAddBatchResultDto> addResult = service.AddFrameObjects(new FrameAddBatchRequestDto
            {
                Frames = frames
            });

            var result = new HoweTrussResultDto
            {
                BayCount = input.BayCount,
                Span = input.Span,
                BaySpacing = baySpacing,
                ChordFrameNames = new List<string>(),
                WebFrameNames = new List<string>(),
                FailureReasons = new List<string>()
            };

            if (!addResult.IsSuccess || addResult.Data == null)
            {
                result.FailureReasons.Add(addResult.Message ?? "Failed to add Howe truss frame objects.");
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
            List<string> chordNames)
        {
            for (int i = 0; i < input.BayCount; i++)
            {
                string bottomName = $"{input.NamePrefix}_BC{i + 1:000}";
                string topName = $"{input.NamePrefix}_TC{i + 1:000}";
                chordNames.Add(bottomName);
                chordNames.Add(topName);

                frames.Add(CreateFrame(bottomName, input.ChordPropName, input.StartX + i * baySpacing, input.StartY, input.StartZ, input.StartX + (i + 1) * baySpacing, input.StartY, input.StartZ));
                frames.Add(CreateFrame(topName, input.ChordPropName, input.StartX + i * baySpacing, input.StartY, input.StartZ + input.Height, input.StartX + (i + 1) * baySpacing, input.StartY, input.StartZ + input.Height));
            }
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
                frames.Add(CreateFrame(name, input.WebPropName, x, input.StartY, input.StartZ, x, input.StartY, input.StartZ + input.Height));
            }
        }

        private static void AddHoweDiagonalFrames(
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

                if (i < center)
                {
                    frames.Add(CreateFrame(
                        name,
                        input.WebPropName,
                        input.StartX + i * baySpacing,
                        input.StartY,
                        input.StartZ + input.Height,
                        input.StartX + (i + 1) * baySpacing,
                        input.StartY,
                        input.StartZ));
                }
                else
                {
                    frames.Add(CreateFrame(
                        name,
                        input.WebPropName,
                        input.StartX + i * baySpacing,
                        input.StartY,
                        input.StartZ,
                        input.StartX + (i + 1) * baySpacing,
                        input.StartY,
                        input.StartZ + input.Height));
                }
            }
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
                NamePrefix = string.IsNullOrWhiteSpace(source.NamePrefix) ? "HT" : source.NamePrefix.Trim(),
                ChordPropName = string.IsNullOrWhiteSpace(source.ChordPropName) ? "Default" : source.ChordPropName.Trim(),
                WebPropName = string.IsNullOrWhiteSpace(source.WebPropName) ? "Default" : source.WebPropName.Trim()
            };
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
