using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OperationLogger
{
    public partial class MainComponent : Component
    {
        //変数の定義
        DateTime _termStartDateTime;        //計測周期のスタート日時
        DateTime _lastOperationDateTime;    //最後に操作した日時
        TimeSpan _judgeTimeSpan;            //無操作の判定時間
        TimeSpan _dayRemainingTime;          //その日の残り時間

        bool _enableKeyboardHook;           //キーボードをフックするか
        bool _enableMouseMoveHook;          //マウスの動きをフックするか
        bool _enableMouseClickHook;         //マウスクリックをフックするか

        const string SETTINGFILE = "Setting.xml";      //設定ファイルの名前
        const string LOGFILE = "OperationLog.csv";     //ログファイルの名前
        const string TEMPLOG = "TempLog.csv";          //一時ログファイルの名前

        public MainComponent()
        {
            //コンポーネントの初期化
            InitializeComponent();

            // TODO:設定データの読み込み

            // 変数の初期化
            _termStartDateTime = DateTime.Now;              //現在時刻を代入
            _lastOperationDateTime = _termStartDateTime;    //開始時はイコール
            _dayRemainingTime = TimeSpan.FromDays(1) - _termStartDateTime.TimeOfDay;    //一日の残り時間


            //設定から値を読み込む
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(SETTINGFILE);
            XmlElement rootElement = xmlDoc.DocumentElement;
            _judgeTimeSpan =
                TimeSpan.FromMinutes(int.Parse(
                    rootElement.GetElementsByTagName("JudgeTimeSpan").Item(0).InnerText));
            _enableKeyboardHook =
                bool.Parse(rootElement.GetElementsByTagName("EnableKeyboardHook").Item(0).InnerText);
            _enableMouseMoveHook =
                bool.Parse(rootElement.GetElementsByTagName("EnableMouseMoveHook").Item(0).InnerText);
            _enableMouseClickHook =
                bool.Parse(rootElement.GetElementsByTagName("EnableMouseClickHook").Item(0).InnerText);

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
        /// 一時ログレコードを作成する
        /// </summary>
        private void makeTempLog()
        {
            //書き込むstringの作成 "PC名,日付,時刻,"
            string hostName = Environment.MachineName;
            string strDate = _termStartDateTime.ToShortDateString();
            string strTime = _termStartDateTime.ToShortTimeString();
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
        private void toolStripMenuItem_Exit_Click(object sender, EventArgs e)
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
            Debug.WriteLine(_termStartDateTime.ToShortDateString());
            //_enableKeyboardHookの値がfalseなら抜ける
            if (_enableKeyboardHook == false) { return; }
            //現在時刻等をログメソッドに送る
            processingOperation(_termStartDateTime, _lastOperationDateTime, DateTime.Now);
            //Debug.WriteLine("Keyboard Logging.");
        }

        /// <summary>
        /// マウス入力をフックしたとき
        /// </summary>
        /// <param name="sender">呼び出し元オブジェクト</param>
        /// <param name="e">イベントデータ</param>
        private void mouseHook_MouseHooked(object sender, MouseHookedEventArgs e)
        {
            //Debug.WriteLine("Mouse Hooked.");
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
            //現在時刻等をログメソッドに送る
            processingOperation(_termStartDateTime, _lastOperationDateTime, DateTime.Now);
            //Debug.WriteLine("Mouse Logging.");
        }

        /// <summary>
        /// 操作日時を受け取って、ログに残すか計算する
        /// </summary>
        /// <param name="datetime">操作日時</param>
        private void processingOperation(
            DateTime termStartDateTime, DateTime lastOperationDateTime, DateTime nowDateTime)
        {
            //前回の操作時間と比較
            if ((nowDateTime - _lastOperationDateTime) > _judgeTimeSpan)
            {   //無操作と判定された場合
                //非同期で一時ログを書く処理
                Task writeLogTask = new Task(() =>
                {
                    //書き出す文字列を作成
                    //操作時間の計算（日をまたぐパターンを考慮）
                    TimeSpan operationTime = lastOperationDateTime - termStartDateTime;
                    writeTempLog(termStartDateTime, operationTime, true);
                    //無操作時間の計算（日をまたぐパターンを考慮）
                    TimeSpan noOperationTime = nowDateTime - lastOperationDateTime;
                    writeTempLog(termStartDateTime, noOperationTime, false);
                });
                writeLogTask.Start();
                _termStartDateTime = nowDateTime;
            }
            //変数の更新
            _lastOperationDateTime = nowDateTime;
        }

        /// <summary>
        /// 操作・無操作時間を受け取り、一時ログに書く
        /// </summary>
        /// <param name="writeTime">操作・無操作時間</param>
        /// <param name="isOperationTime">操作時間=true,無操作時間=false</param>
        public void writeTempLog(DateTime termStartDateTime, TimeSpan writeTime, bool isOperationTime)
        {
            //一時ログファイルを開く
            StreamWriter sw = new StreamWriter(TEMPLOG, true, Encoding.GetEncoding("Shift_JIS"));
            string strWriteTime;    //00:00形式の操作・無操作時間文字列
            string logStr;          //一時ログに書き込む文字列

            int i = 1;  //ループカウンタ
            while (writeTime > _dayRemainingTime)
            {   //その日の残り時間を操作時間が超える場合は繰り返す
                //その日のログは残り時間までで終了、余りを次の日に回す
                strWriteTime =
                    _dayRemainingTime.Hours.ToString() + ":" + _dayRemainingTime.Minutes.ToString();
                //一時ログに書き出す
                logStr = "," + strWriteTime;
                sw.WriteLine(logStr);   //改行あり

                //一時ログに新しい行を作成
                //書き込むstringの作成 "PC名,日付,時刻,"
                string hostName = Environment.MachineName;
                string strDate = termStartDateTime.AddDays(i).ToShortDateString(); //１日加算
                string strTime = "0:0";     //０時起算
                logStr = hostName + "," + strDate + "," + strTime;
                sw.Write(logStr);   //改行なし

                if (isOperationTime == false)
                {   //writeTimeが無操作時間の場合
                    //操作時間"0"を書き込んでおく
                    logStr = ",0:0";
                    sw.Write(logStr);
                }
                //変数の再計算
                writeTime = writeTime - _dayRemainingTime;  //operationtimeから減算
                _dayRemainingTime = TimeSpan.FromDays(1) - TimeSpan.FromSeconds(1);   //_dayRemainingTimeを初期化(23:59)
                i++;    //ループカウンタをインクリメント
            }

            //日またぎ計算後のログを一時ログに書く
            strWriteTime = writeTime.Hours.ToString() + ":" + writeTime.Minutes.ToString();
            _dayRemainingTime = _dayRemainingTime - writeTime;  //_dayRemainingTimeから減算
            logStr = "," + strWriteTime;
            //一時ログファイルに書き出す
            Debug.WriteLine("writeTempLog," + isOperationTime.ToString() + logStr);
            sw.Write(logStr);   //改行しない
            sw.Close();
        }

        /// <summary>
        /// 終了時の処理（ApplicationExitから呼ばれる）
        /// </summary>
        public void writeEndLog()
        {
            string logStr;
            DateTime nowDateTime = DateTime.Now;
            //最後のログを書き込んだ一時ログを本ログファイルに書き出し
            if ((nowDateTime - _lastOperationDateTime) > _judgeTimeSpan)
            {   //無操作と判定された場合
                //書き出す文字列を作成
                TimeSpan operationTime = _lastOperationDateTime - _termStartDateTime;
                string strOperationTime =
                    operationTime.Hours.ToString() + ":" + operationTime.Minutes.ToString();
                TimeSpan noOperationTime = nowDateTime - _lastOperationDateTime;
                string strNoOperationTime =
                    noOperationTime.Hours.ToString() + ":" + noOperationTime.Minutes.ToString();
                logStr = "," + strOperationTime + "," + strNoOperationTime;
            }
            else
            {   //操作ありと判定された場合
                TimeSpan operationTime = nowDateTime - _termStartDateTime;
                string strOperationTime =
                    operationTime.Hours.ToString() + ":" + operationTime.Minutes.ToString();
                logStr = "," + strOperationTime;
            }
            //一時ログファイルにログを書き出し
            Debug.WriteLine("writeEndLog" + logStr);
            StreamWriter sw1 = new StreamWriter(TEMPLOG, true, Encoding.GetEncoding("Shift_JIS"));
            sw1.Write(logStr);  //改行しない
            sw1.Close();
            //一時ログを本ログファイルに書き出し
            StreamReader sr = new StreamReader(TEMPLOG);
            StreamWriter sw2 = new StreamWriter(LOGFILE, true, Encoding.GetEncoding("Shift_JIS"));
            sw2.WriteLine(sr.ReadToEnd());
            sr.Close();
            sw2.Close();
            //一時ファイルを消去
            File.Delete(TEMPLOG);
        }
    }
}
