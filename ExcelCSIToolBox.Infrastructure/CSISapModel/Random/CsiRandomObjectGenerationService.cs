using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.CSISapModel.FrameObject;
using ExcelCSIToolBox.Data.CSISapModel.Random;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Data.Models;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel.Random
{
    public sealed class CsiRandomObjectGenerationService
    {
        private const int DefaultPointCount = 5;
        private const int DefaultFrameCount = 3;
        private const int DefaultShellCount = 1;
        private const int MaxPointCount = 100;
        private const int MaxFrameCount = 100;
        private const int MaxShellCount = 25;

        public OperationResult<RandomCsiObjectResultDto> Generate(
            ICSISapModelConnectionService service,
            RandomCsiObjectRequestDto request)
        {
            if (service == null)
            {
                return OperationResult<RandomCsiObjectResultDto>.Failure("active CSI model is not available.");
            }

            RandomCsiObjectRequestDto normalized = Normalize(request);
            int seed = normalized.Seed ?? Environment.TickCount;
            var rng = new System.Random(seed);
            var result = new RandomCsiObjectResultDto
            {
                Seed = seed,
                PointNames = new List<string>(),
                FrameNames = new List<string>(),
                ShellNames = new List<string>(),
                FailureReasons = new List<string>()
            };

            if (normalized.AddPoints)
            {
                AddRandomPoints(service, normalized, rng, result);
            }

            if (normalized.AddFrames)
            {
                AddRandomFrames(service, normalized, rng, result);
            }

            if (normalized.AddShells)
            {
                AddRandomShells(service, normalized, rng, result);
            }

            result.FailedItems = result.FailureReasons.Count;

            return OperationResult<RandomCsiObjectResultDto>.Success(result);
        }

        private static void AddRandomPoints(
            ICSISapModelConnectionService service,
            RandomCsiObjectRequestDto request,
            System.Random rng,
            RandomCsiObjectResultDto result)
        {
            int count = Clamp(request.PointCount ?? DefaultPointCount, 1, MaxPointCount);
            result.RequestedPoints = count;

            var inputs = new List<CSISapModelPointCartesianInput>();
            for (int i = 0; i < count; i++)
            {
                string name = BuildName(request.PointPrefix, i + 1);
                inputs.Add(new CSISapModelPointCartesianInput
                {
                    ExcelRowNumber = i + 1,
                    UniqueName = name,
                    X = Next(rng, request.MinX.Value, request.MaxX.Value),
                    Y = Next(rng, request.MinY.Value, request.MaxY.Value),
                    Z = Next(rng, request.MinZ.Value, request.MaxZ.Value)
                });
                result.PointNames.Add(name);
            }

            OperationResult<CSISapModelAddPointsResultDTO> addResult = service.AddPointsByCartesian(inputs);
            if (!addResult.IsSuccess || addResult.Data == null)
            {
                result.FailureReasons.Add(addResult.Message ?? "Failed to add random points.");
                result.PointNames.Clear();
                return;
            }

            result.AddedPoints = addResult.Data.AddedCount;
            AddFailures(result, addResult.Data.FailedRowMessages);
        }

        private static void AddRandomFrames(
            ICSISapModelConnectionService service,
            RandomCsiObjectRequestDto request,
            System.Random rng,
            RandomCsiObjectResultDto result)
        {
            int count = Clamp(request.FrameCount ?? DefaultFrameCount, 1, MaxFrameCount);
            result.RequestedFrames = count;

            var frames = new List<FrameAddRequestDto>();
            for (int i = 0; i < count; i++)
            {
                string name = BuildName(request.FramePrefix, i + 1);
                frames.Add(new FrameAddRequestDto
                {
                    UserName = name,
                    PropName = request.FramePropName,
                    Xi = Next(rng, request.MinX.Value, request.MaxX.Value),
                    Yi = Next(rng, request.MinY.Value, request.MaxY.Value),
                    Zi = Next(rng, request.MinZ.Value, request.MaxZ.Value),
                    Xj = Next(rng, request.MinX.Value, request.MaxX.Value),
                    Yj = Next(rng, request.MinY.Value, request.MaxY.Value),
                    Zj = Next(rng, request.MinZ.Value, request.MaxZ.Value)
                });
            }

            OperationResult<FrameAddBatchResultDto> addResult = service.AddFrameObjects(new FrameAddBatchRequestDto
            {
                Frames = frames
            });

            if (!addResult.IsSuccess || addResult.Data == null)
            {
                result.FailureReasons.Add(addResult.Message ?? "Failed to add random frames.");
                return;
            }

            result.AddedFrames = addResult.Data.SuccessCount;
            result.FrameNames.AddRange(addResult.Data.SuccessfulFrameNames ?? new List<string>());
            if (addResult.Data.FailedItems != null)
            {
                foreach (FrameAddResultDto failed in addResult.Data.FailedItems)
                {
                    result.FailureReasons.Add(failed.FailureReason ?? "A random frame failed to add.");
                }
            }
        }

        private static void AddRandomShells(
            ICSISapModelConnectionService service,
            RandomCsiObjectRequestDto request,
            System.Random rng,
            RandomCsiObjectResultDto result)
        {
            int count = Clamp(request.ShellCount ?? DefaultShellCount, 1, MaxShellCount);
            result.RequestedShells = count;

            for (int i = 0; i < count; i++)
            {
                string name = BuildName(request.ShellPrefix, i + 1);
                double x = Next(rng, request.MinX.Value, request.MaxX.Value);
                double y = Next(rng, request.MinY.Value, request.MaxY.Value);
                double z = Next(rng, request.MinZ.Value, request.MaxZ.Value);
                double width = Math.Max((request.MaxX.Value - request.MinX.Value) * 0.08, 100.0);
                double depth = Math.Max((request.MaxY.Value - request.MinY.Value) * 0.08, 100.0);

                var points = new List<CSISapModelShellCoordinateInput>
                {
                    new CSISapModelShellCoordinateInput { X = x, Y = y, Z = z },
                    new CSISapModelShellCoordinateInput { X = Math.Min(x + width, request.MaxX.Value), Y = y, Z = z },
                    new CSISapModelShellCoordinateInput { X = Math.Min(x + width, request.MaxX.Value), Y = Math.Min(y + depth, request.MaxY.Value), Z = z },
                    new CSISapModelShellCoordinateInput { X = x, Y = Math.Min(y + depth, request.MaxY.Value), Z = z }
                };

                OperationResult<string> addResult = service.AddShellByCoord(points, request.ShellPropName, name, "Global", true);
                if (addResult.IsSuccess)
                {
                    result.AddedShells++;
                    result.ShellNames.Add(string.IsNullOrWhiteSpace(addResult.Data) ? name : addResult.Data);
                }
                else
                {
                    result.FailureReasons.Add(addResult.Message ?? "A random shell failed to add.");
                }
            }
        }

        private static RandomCsiObjectRequestDto Normalize(RandomCsiObjectRequestDto request)
        {
            RandomCsiObjectRequestDto source = request ?? new RandomCsiObjectRequestDto();
            bool anyObjectType = source.AddPoints || source.AddFrames || source.AddShells;

            return new RandomCsiObjectRequestDto
            {
                AddPoints = anyObjectType ? source.AddPoints : true,
                AddFrames = source.AddFrames,
                AddShells = source.AddShells,
                PointCount = source.PointCount,
                FrameCount = source.FrameCount,
                ShellCount = source.ShellCount,
                MinX = source.MinX ?? 0,
                MaxX = source.MaxX ?? 10000,
                MinY = source.MinY ?? 0,
                MaxY = source.MaxY ?? 10000,
                MinZ = source.MinZ ?? 0,
                MaxZ = source.MaxZ ?? 3000,
                PointPrefix = string.IsNullOrWhiteSpace(source.PointPrefix) ? "RP" : source.PointPrefix.Trim(),
                FramePrefix = string.IsNullOrWhiteSpace(source.FramePrefix) ? "RF" : source.FramePrefix.Trim(),
                ShellPrefix = string.IsNullOrWhiteSpace(source.ShellPrefix) ? "RS" : source.ShellPrefix.Trim(),
                FramePropName = string.IsNullOrWhiteSpace(source.FramePropName) ? "Default" : source.FramePropName.Trim(),
                ShellPropName = string.IsNullOrWhiteSpace(source.ShellPropName) ? "Default" : source.ShellPropName.Trim(),
                Seed = source.Seed
            };
        }

        private static double Next(System.Random rng, double min, double max)
        {
            if (max < min)
            {
                double swap = min;
                min = max;
                max = swap;
            }

            return min + rng.NextDouble() * (max - min);
        }

        private static int Clamp(int value, int min, int max)
        {
            return Math.Max(min, Math.Min(max, value));
        }

        private static string BuildName(string prefix, int index)
        {
            return (string.IsNullOrWhiteSpace(prefix) ? "R" : prefix.Trim()) +
                   index.ToString("000", CultureInfo.InvariantCulture);
        }

        private static void AddFailures(RandomCsiObjectResultDto result, IReadOnlyList<string> failures)
        {
            if (failures == null)
            {
                return;
            }

            foreach (string failure in failures)
            {
                if (!string.IsNullOrWhiteSpace(failure))
                {
                    result.FailureReasons.Add(failure);
                }
            }
        }
    }
}
