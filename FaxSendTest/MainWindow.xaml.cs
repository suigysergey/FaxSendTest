using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Printing;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using FAXCOMEXLib;

namespace FaxSendTest
{
    public partial class MainWindow : Window
    {
        public byte[] DocumentBytes;
        public ObservableCollection<string> LogEntries { get; } = new ObservableCollection<string>();

        // 用于在事件中传递状态
        private TaskCompletionSource<bool> _faxCompletionSource;

        private FaxServer _currentFaxServer;
        private string _currentJobId;

        public MainWindow()
        {
            InitializeComponent();
            LogListBox.ItemsSource = LogEntries;

            LogEntries.CollectionChanged += (s, e) =>
            {
                if (e.Action == System.Collections.Specialized.NotifyCollectionChangedAction.Add)
                {
                    LogListBox.Dispatcher.BeginInvoke(
                        new Action(() =>
                        {
                            if (LogEntries.Count > 0)
                                LogListBox.ScrollIntoView(LogEntries[^1]);
                        }),
                        System.Windows.Threading.DispatcherPriority.Background
                    );
                }
            };
        }

        private void AddLog(string message)
        {
            Dispatcher.Invoke(() =>
            {
                LogEntries.Add($"[{DateTime.Now:HH:mm:ss}] {message}");
            });
        }

        private string FindFaxPrinterName()
        {
            try
            {
                using var printServer = new LocalPrintServer();
                var queues = printServer.GetPrintQueues(new[] { EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections });

                foreach (var queue in queues)
                {
                    // 方法1：检查打印机名称是否包含 "Fax"
                    if (queue.Name.IndexOf("fax", StringComparison.OrdinalIgnoreCase) >= 0)
                        return queue.Name;

                    // 方法2（更可靠）：检查打印机驱动是否为传真驱动
                    // 传真驱动通常包含 "Fax" 或 "Microsoft Shared Fax Driver"
                    if (queue.QueueDriver?.Name?.IndexOf("Fax", StringComparison.OrdinalIgnoreCase) >= 0)
                        return queue.Name;
                }
            }
            catch (Exception ex)
            {
                AddLog($"查找传真设备失败: {ex.Message}");
            }

            return null;
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(deviceName.Text))
            {
                AddLog("未设置传真设备名");
                return;
            }

            if (string.IsNullOrWhiteSpace(receiveNumber.Text))
            {
                AddLog("未设置接收传真号");
                return;
            }

            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test.pdf");
            if (!File.Exists(path))
            {
                AddLog("传真文档不存在: " + path);
                return;
            }

            try
            {
                DocumentBytes = File.ReadAllBytes(path);
                if (DocumentBytes.Length == 0)
                {
                    AddLog("传真文档为空");
                    return;
                }
            }
            catch (Exception ex)
            {
                AddLog($"读取文档失败: {ex.Message}");
                return;
            }

            AddLog("开始发送传真...");
            try
            {
                bool success = await SendFaxAsync();
                AddLog(success ? "✅ 传真发送成功！" : "❌ 传真发送失败。");
            }
            catch (Exception ex)
            {
                AddLog($"❌ 传真发送异常：{ex.Message}");
            }
        }

        private Task<bool> SendFaxAsync()
        {
            // 确保前一个传真已完成（简单场景，不支持并发）
            if (_currentFaxServer != null)
            {
                AddLog("上一个传真仍在进行中，请稍后再试。");
                return Task.FromResult(false);
            }

            _faxCompletionSource = new TaskCompletionSource<bool>();
            string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".tiff");

            try
            {
                File.WriteAllBytes(tempPath, DocumentBytes);

                _currentFaxServer = new FaxServer();
                _currentFaxServer.Connect("");

                // 注册事件（必须在 Connect 之后）
                RegisterFaxServerEvents(_currentFaxServer);

                var faxDoc = new FaxDocument
                {
                    Body = tempPath,
                    DocumentName = "Pdf文档",
                    Priority = FAX_PRIORITY_TYPE_ENUM.fptLOW
                };
                faxDoc.Sender.Name = "桥景咨询";
                faxDoc.Subject = "Fax Send Test";
                faxDoc.Recipients.Add(receiveNumber.Text, "");
                var jobIds = faxDoc.ConnectedSubmit(_currentFaxServer);
                _currentJobId = jobIds is Array ? jobIds[0] : jobIds;
                if (string.IsNullOrEmpty(_currentJobId))
                {
                    throw new Exception("传真提交失败：未返回作业ID");
                }

                AddLog($"传真作业已提交，ID: {_currentJobId}");
            }
            catch (Exception ex)
            {
                // 清理资源
                CleanupFaxResources(tempPath);
                _faxCompletionSource?.TrySetResult(false);
                throw;
            }

            // 返回任务，由事件完成
            return _faxCompletionSource.Task.ContinueWith(t =>
            {
                CleanupFaxResources(tempPath);
                return t.Result;
            });
        }

        private void RegisterFaxServerEvents(FaxServer faxServer)
        {
            faxServer.OnOutgoingJobAdded += faxServer_OnOutgoingJobAdded;
            faxServer.OnOutgoingJobChanged += faxServer_OnOutgoingJobChanged;
            faxServer.OnOutgoingJobRemoved += faxServer_OnOutgoingJobRemoved;

            var eventsToListen =
                FAX_SERVER_EVENTS_TYPE_ENUM.fsetOUT_QUEUE |
                FAX_SERVER_EVENTS_TYPE_ENUM.fsetQUEUE_STATE |
                FAX_SERVER_EVENTS_TYPE_ENUM.fsetACTIVITY;

            faxServer.ListenToServerEvents(eventsToListen);
        }

        private void faxServer_OnOutgoingJobAdded(FaxServer pFaxServer, string bstrJobId)
        {
            if (bstrJobId == _currentJobId)
            {
                AddLog("传真已加入发送队列");
            }
        }

        private void faxServer_OnOutgoingJobChanged(FaxServer pFaxServer, string bstrJobId, FaxJobStatus pJobStatus)
        {
            if (bstrJobId != _currentJobId) return;

            pFaxServer.Folders.OutgoingQueue.Refresh();
            string statusText = GetJobStatusText(pJobStatus.Status, pJobStatus.ExtendedStatusCode);
            AddLog($"作业状态更新: {statusText}");

            switch (pJobStatus.Status)
            {
                case FAX_JOB_STATUS_ENUM.fjsCOMPLETED:
                    _faxCompletionSource?.TrySetResult(true);
                    break;

                case FAX_JOB_STATUS_ENUM.fjsFAILED:
                case FAX_JOB_STATUS_ENUM.fjsRETRIES_EXCEEDED:
                case FAX_JOB_STATUS_ENUM.fjsCANCELED:
                    _faxCompletionSource?.TrySetResult(false);
                    break;

                    // 其他状态继续等待
            }
        }

        private void faxServer_OnOutgoingJobRemoved(FaxServer pFaxServer, string bstrJobId)
        {
            if (bstrJobId != _currentJobId) return;

            AddLog("传真作业已从队列移除");
            // 如果之前未设置结果，可能表示成功（但保守起见，仅当 COMPLETED 时才算成功）
            // 此处不自动设为成功，依赖 OnOutgoingJobChanged 中的 fjsCOMPLETED
        }

        private string GetJobStatusText(FAX_JOB_STATUS_ENUM status, FAX_JOB_EXTENDED_STATUS_ENUM extended)
        {
            string baseStatus = status switch
            {
                FAX_JOB_STATUS_ENUM.fjsPENDING => "等待中",
                FAX_JOB_STATUS_ENUM.fjsINPROGRESS => "进行中",
                FAX_JOB_STATUS_ENUM.fjsFAILED => "失败",
                FAX_JOB_STATUS_ENUM.fjsRETRIES_EXCEEDED => "重试超限",
                FAX_JOB_STATUS_ENUM.fjsCOMPLETED => "已完成",
                FAX_JOB_STATUS_ENUM.fjsCANCELED => "已取消",
                FAX_JOB_STATUS_ENUM.fjsRETRYING => "重试中",
                FAX_JOB_STATUS_ENUM.fjsROUTING => "路由中",
                FAX_JOB_STATUS_ENUM.fjsPAUSED => "已暂停",
                FAX_JOB_STATUS_ENUM.fjsNOLINE => "无线路",
                FAX_JOB_STATUS_ENUM.fjsCANCELING => "取消中",
                _ => status.ToString()
            };

            string extStatus = extended switch
            {
                FAX_JOB_EXTENDED_STATUS_ENUM.fjesDIALING => "拨号中",
                FAX_JOB_EXTENDED_STATUS_ENUM.fjesTRANSMITTING => "发送中",
                FAX_JOB_EXTENDED_STATUS_ENUM.fjesCALL_COMPLETED => "呼叫完成",
                _ => ""
            };

            return string.IsNullOrEmpty(extStatus) ? baseStatus : $"{baseStatus} ({extStatus})";
        }

        private void CleanupFaxResources(string tempPath)
        {
            try
            {
                // 取消事件订阅
                if (_currentFaxServer != null)
                {
                    _currentFaxServer.OnOutgoingJobAdded -= faxServer_OnOutgoingJobAdded;
                    _currentFaxServer.OnOutgoingJobChanged -= faxServer_OnOutgoingJobChanged;
                    _currentFaxServer.OnOutgoingJobRemoved -= faxServer_OnOutgoingJobRemoved;
                    _currentFaxServer.Disconnect();
                    Marshal.ReleaseComObject(_currentFaxServer);
                    _currentFaxServer = null;
                }

                if (File.Exists(tempPath))
                    File.Delete(tempPath);

                _currentJobId = null;
            }
            catch (Exception ex)
            {
                AddLog($"清理资源时出错: {ex.Message}");
            }
        }

        private void GetDefaultFaxName(object sender, RoutedEventArgs e)
        {
            this.deviceName.Text = FindFaxPrinterName();
        }
    }
}