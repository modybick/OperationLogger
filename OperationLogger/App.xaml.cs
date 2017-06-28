using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace OperationLogger
{
    /// <summary>
    /// App.xaml の相互作用ロジック
    /// </summary>
    public partial class App : Application
    {
        /// <summary>
        /// タスクトレイに表示するアイコン
        /// </summary>
        private MainComponent mainComponent;

        /// <summary>
        /// System.Windows.Application.Startupイベントを発生させる
        /// </summary>
        /// <param name="e">イベントデータを格納しているStartupEventArgs</param>
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            this.ShutdownMode = ShutdownMode.OnExplicitShutdown;
            this.mainComponent = new MainComponent();
        }

        /// <summary>
        /// System.Windows.Application.Exitイベントを発生させる
        /// </summary>
        /// <param name="e">イベントデータを格納しているExitEventArgs</param>
        protected override void OnExit(ExitEventArgs e)
        {
            //終了時の処理
            this.mainComponent.endProcess();
            base.OnExit(e);
            this.mainComponent.Dispose();
        }
    }
}
