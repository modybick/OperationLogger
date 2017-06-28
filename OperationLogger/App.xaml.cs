using System.Windows;

namespace OperationLogger
{
    /// <summary>
    /// App.xaml の相互作用ロジック
    /// </summary>
    public partial class App : Application
    {
        System.Threading.Mutex _mutex;

        /// <summary>
        /// メインコンポーネント
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
            base.OnExit(e);
            this.mainComponent.Dispose();
        }

        private void Application_Startup(object sender, StartupEventArgs e)
        {
            //２重起動を防止

            //Mutex名を決める
            string mutexName = "OperationLogger";
            mutexName = "Global\\" + mutexName;
            //すべてのユーザーにフルコントロールを許可するMutexSecurityを作成する
            System.Security.AccessControl.MutexAccessRule rule =
                new System.Security.AccessControl.MutexAccessRule(
                    new System.Security.Principal.SecurityIdentifier(
                        System.Security.Principal.WellKnownSidType.WorldSid, null),
                    System.Security.AccessControl.MutexRights.FullControl,
                    System.Security.AccessControl.AccessControlType.Allow);
            System.Security.AccessControl.MutexSecurity mutexSecurity =
                new System.Security.AccessControl.MutexSecurity();
            mutexSecurity.AddAccessRule(rule);
            //Mutexオブジェクトを作成する
            bool createdNew;
            _mutex =
                new System.Threading.Mutex(false, mutexName, out createdNew, mutexSecurity);

            //mutexの所有権を要求
            if (_mutex.WaitOne(0,false) == false)
            {   //所有権をもらえない場合は既に起動していると判断して終了
                _mutex.ReleaseMutex();
                _mutex.Close();
                return;
            }
        }

        private void Application_Exit(object sender, ExitEventArgs e)
        {
            //終了時の処理
            this.mainComponent.endProcess();

            _mutex.ReleaseMutex();
            _mutex.Close();
        }
    }
}
