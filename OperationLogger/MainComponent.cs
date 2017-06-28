using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace OperationLogger
{
    public partial class MainComponent : Component
    {
        //変数の定義
        DateTime _startDatetime;         //アプリの起動日時
        DateTime _termStartDatetime;     //計測周期のスタート日時
        DateTime _lastOperationDatetime; //最後に操作した日時
        TimeSpan _judgeTimeSpan;         //無操作の判定時間

        bool _enableKeyboardHook;     //キーボードをフックするか
        bool _enableMouseMoveHook;    //マウスの動きをフックするか
        bool _enableMouseClickHook;       //マウスクリックをフックするか

        static string SETTINGFILE = "Setting.ini";      //設定ファイルの名前
        static string LOGFILE = "OperationLog.csv";     //ログファイルの名前
        static string TEMPLOG = "TempLog.csv";          //一時ログファイルの名前

        public MainComponent()
        {
            //コンポーネントの初期化
            InitializeComponent();

            //コンテキストメニューのイベントを設定
            this.toolStripMenuItem_Exit.Click += ToolStripMenuItem_Exit_Click;

            // TODO:設定データの読み込み

            // 変数の初期化
            _startDatetime = DateTime.Now;           //現在時刻を代入
            _termStartDatetime = _startDatetime;      //開始時はイコール
            _lastOperationDatetime = _startDatetime;  //開始時はイコール
            // TODO:設定から読み込むように変更する
            _judgeTimeSpan = TimeSpan.FromMinutes(1);   //ひとまず1分で無操作判定
            _enableKeyboardHook = false;
            _enableMouseMoveHook = true;
            _enableMouseClickHook = false;

            //前回異常終了していないか調べ,異常終了している場合は対処する
            dealAbEnd();
            //一時ログファイルの作成
            makeTempLog();

        }

        /// <summary>
        /// 前回異常終了していないか調べ、している場合は対処する
        /// </summary>
        private void dealAbEnd()
        {
            if (File.Exists(TEMPLOG))
            {   //前回異常終了していると判断

                //残っている一時ログに異常終了を示す値を付加して本ログに書き込む

                string logStr;  //書き込むレコード
                //一時ログファイルの読み込み
                StreamReader sr = new StreamReader(TEMPLOG, Encoding.GetEncoding("Shift_JIS"));
                logStr = sr.ReadLine() + ",ABEND";
                sr.Close();

                //本ログファイルへの書き込み（追加で書き込む）
                StreamWriter sw = new StreamWriter(LOGFILE, true, Encoding.GetEncoding("Shift_JIS"));
                sw.WriteLine(logStr);
                sw.Close();

                //一時ログの削除
                File.Delete(TEMPLOG);
            }

        }

        /// <summary>
        /// 一時ログファイルを作成する
        /// 最初に書き込むべきログは書き込んでおく
        /// </summary>
        private void makeTempLog()
        {
            //書き込むstringの作成 "PC名,日付,時刻,"
            string hostName = Environment.MachineName;
            string strDate = _startDatetime.ToShortDateString();
            string strTime = _startDatetime.ToShortTimeString();
            string logStr = hostName + "," + strDate + "," + strTime;

            //ファイルの作成・書き込み
            StreamWriter sw = new StreamWriter(TEMPLOG, false, Encoding.GetEncoding("Shift_JIS"));
            sw.Write(logStr);   //改行はしない
            sw.Close();
        }

        /// <summary>
        /// コンテキストメニュー”終了”をクリックしたとき
        /// </summary>
        /// <param name="sender">呼び出し元オブジェクト</param>
        /// <param name="e">イベントデータ</param>
        private void ToolStripMenuItem_Exit_Click(object sender, EventArgs e)
        {
            //現在のアプリケーションの終了
            System.Windows.Application.Current.Shutdown();
        }

        /// <summary>
        /// キーボード入力をフックしたとき
        /// </summary>
        /// <param name="sender">呼び出し元オブジェクト</param>
        /// <param name="e">イベントデータ</param>
        private void keyboardHook_KeyboardHooked(object sender, KeyboardHookedEventArgs e)
        {
            Debug.WriteLine("Keyboard Hooked.");
            //_enableKeyboardHookの値がfalseなら抜ける
            if (_enableKeyboardHook == false) { return; }
            // TODO:設定でON/OFFできるようにする
            //現在時刻等をログメソッドに送る
            processingOperation(_termStartDatetime, _lastOperationDatetime, DateTime.Now);
            Debug.WriteLine("Keyboard Logging.");
        }

        /// <summary>
        /// マウス入力をフックしたとき
        /// </summary>
        /// <param name="sender">呼び出し元オブジェクト</param>
        /// <param name="e">イベントデータ</param>
        private void mouseHook_MouseHooked(object sender, MouseHookedEventArgs e)
        {
            Debug.WriteLine("Mouse Hooked.");
            if (e.Message == MouseMessage.Move)
            {   //マウスムーブの場合
                //_enableMouseMoveHookの値がfalseであれば抜ける
                if (_enableMouseMoveHook == false) { return; }
            }
            else
            {   //マウスムーブ以外の場合
                //_enableMouseClickHookの値がfalseであれば抜ける
                if (_enableMouseClickHook == false) { return; }
            }
            // TODO:設定でON/OFFできるようにする
            //現在時刻等をログメソッドに送る
            processingOperation(_termStartDatetime, _lastOperationDatetime, DateTime.Now);
            Debug.WriteLine("Mouse Logging.");
        }

        /// <summary>
        /// 操作日時を受け取って、ログに残すか計算する
        /// </summary>
        /// <param name="datetime">操作日時</param>
        private void processingOperation(
            DateTime termStartDatetime, DateTime lastOperationDatetime, DateTime nowDatetime)
        {
            //前回の操作時間と比較
            if ((nowDatetime - _lastOperationDatetime) > _judgeTimeSpan)
            {   //無操作と判定された場合
                //非同期で一時ログを書く処理
                Task writeLogTask = new Task(() =>
                {
                    //書き出す文字列を作成
                    TimeSpan operationTime = lastOperationDatetime - termStartDatetime;
                    string strOperationTime =
                        operationTime.Hours.ToString() + ":" + operationTime.Minutes.ToString();
                    TimeSpan noOperationTime = nowDatetime - lastOperationDatetime;
                    string strNoOperationTime =
                        noOperationTime.Hours.ToString() + ":" + noOperationTime.Minutes.ToString();
                    string logStr = "," + strOperationTime + "," + strNoOperationTime;
                    //一時ログファイルに書き出す
                    StreamWriter sw = new StreamWriter(TEMPLOG, true, Encoding.GetEncoding("Shift_JIS"));
                    sw.Write(logStr);
                    sw.Close();
                });
                writeLogTask.Start();
                _termStartDatetime = nowDatetime;
            }
            //変数の更新
            _lastOperationDatetime = nowDatetime;
        }

        /// <summary>
        /// 終了時の処理（ApplicationExitから呼ばれる）
        /// </summary>
        public void endProcess()
        {
            string logStr;
            DateTime nowDatetime = DateTime.Now;
            //最後のログを書き込んだ一時ログを本ログファイルに書き出し
            if ((nowDatetime - _lastOperationDatetime) > _judgeTimeSpan)
            {   //無操作と判定された場合
                //書き出す文字列を作成
                TimeSpan operationTime = _lastOperationDatetime - _termStartDatetime;
                string strOperationTime =
                    operationTime.Hours.ToString() + ":" + operationTime.Minutes.ToString();
                TimeSpan noOperationTime = nowDatetime - _lastOperationDatetime;
                string strNoOperationTime =
                    noOperationTime.Hours.ToString() + ":" + noOperationTime.Minutes.ToString();
                logStr = "," + strOperationTime + "," + strNoOperationTime;
            }
            else
            {   //操作ありと判定された場合
                TimeSpan operationTime = nowDatetime - _termStartDatetime;
                string strOperationTime =
                    operationTime.Hours.ToString() + ":" + operationTime.Minutes.ToString();
                logStr = "," + strOperationTime;
            }
            //一時ログファイルにログを書き出し
            StreamWriter sw1 = new StreamWriter(TEMPLOG, true, Encoding.GetEncoding("Shift_JIS"));
            sw1.Write(logStr);
            sw1.Close();
            //一時ログを本ログファイルに書き出し
            StreamReader sr = new StreamReader(TEMPLOG);
            StreamWriter sw2 = new StreamWriter(LOGFILE, true, Encoding.GetEncoding("Shift_JIS"));
            sw2.WriteLine(sr.ReadLine());
            sr.Close();
            sw2.Close();
            //一時ファイルを消去
            File.Delete(TEMPLOG);
        }

    }
}
