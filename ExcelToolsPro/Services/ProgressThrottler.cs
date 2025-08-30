using System;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace ExcelToolsPro.Services
{
    public class ProgressThrottler(IProgress<float>? progress = null, int minIntervalMs = 100, float minStep = 0.5f, ILogger? logger = null)
    {
        private float _lastValue = -1f;
        private DateTime _lastReportTime = DateTime.MinValue;
        private readonly IProgress<float>? _progress = progress;
        private readonly int _minIntervalMs = minIntervalMs;
        private readonly float _minStep = minStep;
        private readonly ILogger? _logger = logger;
        private int _errorCount = 0;
        private const int MaxErrorCount = 5; // 最大错误次数，超过后停止报告

        public ProgressThrottler() : this(null, 100, 0.5f, null)
        {
        }

        public void Report(float value)
        {
            if (_progress == null || value < 0 || _errorCount >= MaxErrorCount) return;

            try
            {
                var now = DateTime.UtcNow;
                // 确保进度值在0-100之间
                var clampedValue = Math.Clamp(value, 0, 100);

                // 首次报告、达到最小步长、达到最小时间间隔或完成时（100%）
                if (_lastValue < 0f || 
                    clampedValue >= _lastValue + _minStep || 
                    (now - _lastReportTime).TotalMilliseconds >= _minIntervalMs ||
                    Math.Abs(clampedValue - 100f) < 0.001f)
                {
                    _lastValue = clampedValue;
                    _lastReportTime = now;
                    _progress.Report(clampedValue);
                }
            }
            catch (Exception ex)
            {
                _errorCount++;
                _logger?.LogWarning(ex, "进度报告时发生异常，错误次数: {ErrorCount}/{MaxErrorCount}, 进度值: {Value}", 
                    _errorCount, MaxErrorCount, value);
                
                if (_errorCount >= MaxErrorCount)
                {
                    _logger?.LogError("进度报告错误次数超过限制 ({MaxErrorCount})，停止进度报告功能", MaxErrorCount);
                }
            }
        }

        public void Report(float value, bool force)
        {
            if (force)
            {
                if (_progress == null || _errorCount >= MaxErrorCount) return;
                
                try
                {
                    var clampedValue = Math.Clamp(value, 0, 100);
                    _lastValue = clampedValue;
                    _lastReportTime = DateTime.UtcNow;
                    _progress.Report(clampedValue);
                }
                catch (Exception ex)
                {
                    _errorCount++;
                    _logger?.LogWarning(ex, "强制进度报告时发生异常，错误次数: {ErrorCount}/{MaxErrorCount}, 进度值: {Value}", 
                        _errorCount, MaxErrorCount, value);
                    
                    if (_errorCount >= MaxErrorCount)
                    {
                        _logger?.LogError("进度报告错误次数超过限制 ({MaxErrorCount})，停止进度报告功能", MaxErrorCount);
                    }
                }
            }
            else
            {
                Report(value);
            }
        }
        
        /// <summary>
        /// 重置错误计数器
        /// </summary>
        public void ResetErrorCount()
        {
            _errorCount = 0;
            _logger?.LogDebug("进度报告错误计数器已重置");
        }
        
        /// <summary>
        /// 获取当前错误状态
        /// </summary>
        public bool IsErrorLimitReached => _errorCount >= MaxErrorCount;
        
        /// <summary>
        /// 获取当前错误次数
        /// </summary>
        public int ErrorCount => _errorCount;
    }
}