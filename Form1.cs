using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Collections;
using System.Diagnostics;

namespace ExcuteAssistant
{
    public partial class Form1 : Form
    {

		[System.Runtime.InteropServices.DllImport("user32.dll",
                CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern int GetWindowText(int hWnd,
		    StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetClassName(IntPtr hWnd,
	        StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);

		[DllImport("user32.dll", CharSet = CharSet.Auto)]
		static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, int wParam, StringBuilder lParam);

		[DllImport("user32.dll", SetLastError = true)]
		static extern int GetWindow(int hWnd, int uCmd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        public const uint WM_SYSKEYDOWN = 0x0104;
		public const uint WM_SYSKEYUP = 0x0105;
        public const uint WM_KEYDOWN = 0x0100;
        public const uint WM_KEYUP = 0x0101;
		public const uint WM_SETTEXT = 0x000C;
		public const uint WM_CLOSE = 0x0010;

		public const uint SC_CLOSE = 0xF060;
		public const uint WM_SYSCOMMAND = 0x0112;

        public const int GW_HWNDNEXT = 2;
        public const int GW_HWNDPREV = 3;
        public const int GW_CHILD = 5;

        public const uint VK_MENU = 0x0012;
        public const uint VK_LMENU = 0x00A4;
        public const uint VK_RETURN = 0x000D;

        public const uint VK_B = 0x0042;
        public const uint VK_C = 0x0043;
        public const uint VK_T = 0x0054;
        public const uint VK_O = 0x004F;
        public const uint VK_P = 0x0050;
        public const uint VK_R = 0x0052;
        public const uint VK_F = 0x0046;
		public const uint VK_F10 = 0x79;
		public const uint VK_UP = 0x26;
        public const uint VK_DOWN = 0x28;
        public const uint VK_TAB = 0x09;
        public const uint VK_SPACE = 0x20;
        public const uint VK_LEFT = 0x25;
        public const uint VK_RIGHT = 0x27;
        public const uint VK_ESCAPE = 0x1B;

        public int holdHandle = 0;//「メンテナンス画面」
		public int browseHandle = 0;//「フォルダーの参照」

        public List<string> prev_list;
        public int restart_timer_count = 0;
        public bool isRestarting = false;
        public bool isExcuting = false;
		public string[] subFolders;
		public int index = 0xFFFF;

#if false
		string path = "C:\\DocKop\\Epson移載機ソフト関連\\projects";
#else
		string path = "C:\\Users\\kyocera\\Desktop\\projects";
#endif

        public Form1()
        {
            InitializeComponent();
        	label1.Visible = true;
        	label2.Visible = true;

			//フォームが最大化されないようにする
			this.MaximizeBox = false;
			timer1.Enabled = true;
            prev_list = new List<string>();

			this.ActiveControl = this.listView1;
            listView1.FullRowSelect = true;
            listView1.GridLines = true;
            listView1.Sorting = SortOrder.Ascending;
            listView1.View = View.Details;
            listView1.HideSelection = false;
            listView1.ForeColor = Color.Black;//初期の色
            listView1.BackColor = Color.Lime;//全体背景色

            ColumnHeader columnFolder;
            ColumnHeader columnNumber;
			columnFolder = new ColumnHeader();
			columnNumber = new ColumnHeader();
			columnFolder.Text = "型番";
			columnNumber.Text = "No";
            ColumnHeader[] colHeaderRegValue = {columnFolder, columnNumber};
            listView1.Columns.AddRange(colHeaderRegValue);
			listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
        }

		private void SequenceControl(int offset)
		{
			//前の画面が表示されていたら、閉じる
		    StringBuilder sb = new StringBuilder("");
            if (browseHandle != 0)
            {
				SendMessage((IntPtr)browseHandle, WM_SYSCOMMAND, (int)SC_CLOSE, sb); 
	            Thread.Sleep(200);
            }
            if (holdHandle != 0)
			{
				SendMessage((IntPtr)holdHandle, WM_SYSCOMMAND, (int)SC_CLOSE, sb); 
	            Thread.Sleep(200);
			}

			//プロセスよりEpson RC+のメインハンドル取得
			int destWnd = 0;
            System.Diagnostics.Process[] process;
            process = System.Diagnostics.Process.GetProcessesByName("erc70");
            foreach (System.Diagnostics.Process ps in process)
            {
                destWnd = (int)ps.MainWindowHandle;//ターゲットのメインハンドル
                SetForegroundWindow(ps.MainWindowHandle);//ターゲットにFocusを当てる：非アクティブだとメニューにフォーカスが当たらない
            }
            if(destWnd == 0)
            {
				isExcuting = false;
				return;//Epson RC+が起動していない。
			}

			//メニュー→ツール→メンテナンスをショートカットで遷移する
            PostMessage((IntPtr)destWnd, WM_SYSKEYDOWN, (int)VK_MENU, 0);//ALTキーを押下
            Thread.Sleep(200);
            PostMessage((IntPtr)destWnd, WM_SYSKEYUP, (int)VK_MENU, 0);//ALTキーを押下
            Thread.Sleep(200);

            PostMessage((IntPtr)destWnd, WM_KEYDOWN, (int)VK_T, 0);//Tキーを押下
            Thread.Sleep(200);
            PostMessage((IntPtr)destWnd, WM_KEYDOWN, (int)VK_O, 0);//Oキーを押下
            Thread.Sleep(200);

			//メンテナンス画面検索
            int hPrev = GetWindow(destWnd, GW_HWNDPREV);
			do
			{
	            //ウィンドウのタイトルを取得する
	            StringBuilder tsb = new StringBuilder(256);
	            GetWindowText(hPrev, tsb, tsb.Capacity);
                string tsbStr = tsb.ToString();

                //ウィンドウのクラス名を取得する
                StringBuilder csb = new StringBuilder(256);
                GetClassName((IntPtr)hPrev, csb, csb.Capacity);
                string csbStr = csb.ToString();

				label1.Text += csbStr + "\r\n";
				label2.Text += tsbStr + "\r\n";

                if (tsb.ToString() == "メンテナンス")
                {
                    break;
                }

                hPrev = GetWindow(hPrev, GW_HWNDPREV);
		    }
		    while(hPrev != 0);

			holdHandle = hPrev;//メンテナンス・・・次の画面検索の為に保持しておく
            hPrev = GetWindow(hPrev, GW_CHILD);


			//リストアボタン検索
            hPrev = GetWindow(hPrev, GW_CHILD);
			do
			{
	            //ウィンドウのタイトルを取得する
	            StringBuilder tsb = new StringBuilder(256);
	            GetWindowText(hPrev, tsb, tsb.Capacity);
                string tsbStr = tsb.ToString();

                //ウィンドウのクラス名を取得する
                StringBuilder csb = new StringBuilder(256);
                GetClassName((IntPtr)hPrev, csb, csb.Capacity);
                string csbStr = csb.ToString();

				label1.Text += csbStr + "\r\n";
				label2.Text += tsbStr + "\r\n";

                if (tsbStr == "コントローラー設定リストア(&R)...")
                {
//					label2.Text += "HIT" + "\r\n";
                    break;
                }

                hPrev = GetWindow(hPrev, GW_HWNDNEXT);
		    }
		    while(hPrev != 0);

			Thread.Sleep(200);
#if false
			PostMessage((IntPtr)hPrev, WM_KEYDOWN, (int)VK_R, 0);//Rキーを押下(リストア)//ショートカット押下だと、バックアップボタンにフォーカスが残る為NG
#else
			PostMessage((IntPtr)hPrev, WM_KEYDOWN, (int)VK_TAB, 0);//TABキーを押下
			Thread.Sleep(200);
			PostMessage((IntPtr)hPrev, WM_KEYUP, (int)VK_TAB, 0);//TABキーを戻す
			Thread.Sleep(200);
			PostMessage((IntPtr)hPrev, WM_KEYDOWN, (int)VK_SPACE, 0);//SPACEキーを押下
			Thread.Sleep(200);
			PostMessage((IntPtr)hPrev, WM_KEYUP, (int)VK_SPACE, 0);//SPACEキーを戻す
#endif
			Thread.Sleep(200);


            hPrev = GetWindow(holdHandle, GW_HWNDPREV);//「フォルダーの参照」を取得
			browseHandle = hPrev;
            hPrev = GetWindow(hPrev, GW_CHILD);//画面から隠れているEditBoxを取得
            int editInst = 0;
            int browserInst = 0;
			do
			{
	            //ウィンドウのタイトルを取得する
	            StringBuilder tsb = new StringBuilder(256);
	            GetWindowText(hPrev, tsb, tsb.Capacity);
                string tsbStr = tsb.ToString();

                //ウィンドウのクラス名を取得する
                StringBuilder csb = new StringBuilder(256);
                GetClassName((IntPtr)hPrev, csb, csb.Capacity);
                string csbStr = csb.ToString();

				label1.Text += csbStr + "\r\n";
				label2.Text += tsbStr + "\r\n";

                if(csbStr == "Edit")
                {
					editInst = hPrev;
//					label1.Text += "Edit HIT" + "\r\n";
				}

                if(csbStr == "SHBrowseForFolder ShellNameSpace Control")
                {
					browserInst = hPrev;
//					label1.Text += "Browser HIT" + "\r\n";
				}

                hPrev = GetWindow(hPrev, GW_HWNDNEXT);
		    }
		    while(hPrev != 0);

			//ツリーの中でフォルダ移動
            int treeInst = 0;
            hPrev = GetWindow(browserInst, GW_CHILD);//画面から隠れているTreeを取得
			do
			{
	            //ウィンドウのタイトルを取得する
	            StringBuilder tsb = new StringBuilder(256);
	            GetWindowText(hPrev, tsb, tsb.Capacity);
                string tsbStr = tsb.ToString();

                //ウィンドウのクラス名を取得する
                StringBuilder csb = new StringBuilder(256);
                GetClassName((IntPtr)hPrev, csb, csb.Capacity);
                string csbStr = csb.ToString();

				label1.Text += csbStr + "\r\n";
				label2.Text += tsbStr + "\r\n";

                if(tsbStr == "ツリー表示")
                {
					treeInst = hPrev;
//					label1.Text += "Tree HIT" + "\r\n";
				}

                hPrev = GetWindow(hPrev, GW_HWNDNEXT);
		    }
		    while(hPrev != 0);

			Thread.Sleep(500);

			PostMessage((IntPtr)treeInst, WM_KEYDOWN, (int)VK_SPACE, 0);//SPACEキーを押下(フォーカスを与える)
			Thread.Sleep(50);
			PostMessage((IntPtr)treeInst, WM_KEYUP, (int)VK_SPACE, 0);//SPACEキーを戻す
			Thread.Sleep(50);

			for(int i = 0; i < 5; i++)//繰り返しで「C:\」までカーソル移動
			{
				PostMessage((IntPtr)treeInst, WM_KEYDOWN, (int)VK_LEFT, 0);//左キーを押下
				Thread.Sleep(50);
				PostMessage((IntPtr)treeInst, WM_KEYUP, (int)VK_LEFT, 0);//左キーを戻す
				Thread.Sleep(50);
			}

			for(int i = 0; i < 3; i++)//繰り返しで「デスクトップ」までカーソル移動
			{
				PostMessage((IntPtr)treeInst, WM_KEYDOWN, (int)VK_DOWN, 0);//下キーを押下
				Thread.Sleep(50);
				PostMessage((IntPtr)treeInst, WM_KEYUP, (int)VK_DOWN, 0);//下キーを戻す
				Thread.Sleep(50);
			}

			for(int i = 0; i < 2; i++)//繰り返しで「projects」までカーソル移動
			{
				PostMessage((IntPtr)treeInst, WM_KEYDOWN, (int)VK_RIGHT, 0);//右キーを押下
				Thread.Sleep(50);
				PostMessage((IntPtr)treeInst, WM_KEYUP, (int)VK_RIGHT, 0);//右キーを戻す
				Thread.Sleep(50);
			}

			//ここより機種別
			if(offset > 0)
			{
				for(int i = 0; i < offset; i++)//各機種までカーソル移動
				{
					PostMessage((IntPtr)treeInst, WM_KEYDOWN, (int)VK_DOWN, 0);//下キーを押下
					Thread.Sleep(50);
					PostMessage((IntPtr)treeInst, WM_KEYUP, (int)VK_DOWN, 0);//下キーを戻す
					Thread.Sleep(50);
				}

				//右キー
				PostMessage((IntPtr)treeInst, WM_KEYDOWN, (int)VK_RIGHT, 0);//右キーを押下
				Thread.Sleep(50);
				PostMessage((IntPtr)treeInst, WM_KEYUP, (int)VK_RIGHT, 0);//右キーを戻す
				Thread.Sleep(50);
				//下キー
				PostMessage((IntPtr)treeInst, WM_KEYDOWN, (int)VK_DOWN, 0);//下キーを押下
				Thread.Sleep(50);
				PostMessage((IntPtr)treeInst, WM_KEYUP, (int)VK_DOWN, 0);//下キーを戻す
				Thread.Sleep(50);

				//Enterキー
		        PostMessage((IntPtr)editInst, WM_KEYDOWN, (int)VK_RETURN, 0);//Enterキーを押下
		        Thread.Sleep(50);
				PostMessage((IntPtr)editInst, WM_KEYUP, (int)VK_RETURN, 0);//Enterキーを戻す
				Thread.Sleep(50);


                //「コントローラーリストア」画面を探す
	            Thread.Sleep(500);
	            IntPtr hWnd = FindWindow(null, "コントローラーリストア");
				if (hWnd != null)
				{
				    //ウィンドウタイトルを取得
				    StringBuilder title = new StringBuilder(256);
				    int titleLen = GetWindowText((int)hWnd, title, 256);
					label1.Text += title + "HIT" + "\r\n";


					//「プロジェクト」のみチェックを入れる
			        Thread.Sleep(100);
			        PostMessage((IntPtr)hWnd, WM_KEYDOWN, (int)VK_P, 0);//Pキーを押下
//			        Thread.Sleep(100);
//					PostMessage((IntPtr)hWnd, WM_KEYUP, (int)VK_P, 0);//Pキーを戻す //チェックボックスはWM_KEYUPで戻ってしまう為コメントアウト、またSPACEキーでチェックされず。

					//OKまでスキップ
			        Thread.Sleep(100);
			        PostMessage((IntPtr)hWnd, WM_KEYDOWN, (int)VK_TAB, 0);//Tabキーを押下
			        Thread.Sleep(100);
					PostMessage((IntPtr)hWnd, WM_KEYUP, (int)VK_TAB, 0);//Tabキーを戻す
					Thread.Sleep(100);
			        PostMessage((IntPtr)hWnd, WM_KEYDOWN, (int)VK_TAB, 0);//Tabキーを押下
			        Thread.Sleep(100);
					PostMessage((IntPtr)hWnd, WM_KEYUP, (int)VK_TAB, 0);//Tabキーを戻す
		            Thread.Sleep(100);

					//Enterキー
			        PostMessage((IntPtr)hWnd, WM_KEYDOWN, (int)VK_RETURN, 0);//Enterキーを押下
			        Thread.Sleep(100);
					PostMessage((IntPtr)hWnd, WM_KEYUP, (int)VK_RETURN, 0);//Enterキーを戻す
					Thread.Sleep(100);


#if false//debug用
		            int hwndwnd = GetWindow((int)hWnd, GW_CHILD);/////
					do
					{
			            //ウィンドウのタイトルを取得する
			            StringBuilder tsb = new StringBuilder(256);
			            GetWindowText(hwndwnd, tsb, tsb.Capacity);
		                string tsbStr = tsb.ToString();

		                //ウィンドウのクラス名を取得する
		                StringBuilder csb = new StringBuilder(256);
		                GetClassName((IntPtr)hwndwnd, csb, csb.Capacity);
		                string csbStr = csb.ToString();

						label1.Text += csbStr + "\r\n";
						label2.Text += tsbStr + "\r\n";

		                if(tsbStr == "")
		                {
							break;
						}

		                hwndwnd = GetWindow(hwndwnd, GW_HWNDNEXT);
				    }
				    while(hwndwnd != 0);

		            hwndwnd = GetWindow((int)hwndwnd, GW_CHILD);/////
					do
					{
			            //ウィンドウのタイトルを取得する
			            StringBuilder tsb = new StringBuilder(256);
			            GetWindowText(hwndwnd, tsb, tsb.Capacity);
		                string tsbStr = tsb.ToString();

		                //ウィンドウのクラス名を取得する
		                StringBuilder csb = new StringBuilder(256);
		                GetClassName((IntPtr)hwndwnd, csb, csb.Capacity);
		                string csbStr = csb.ToString();

						label1.Text += csbStr + "\r\n";
						label2.Text += tsbStr + "\r\n";

		                hwndwnd = GetWindow(hwndwnd, GW_HWNDNEXT);
				    }
				    while(hwndwnd != 0);
#endif
				}

				restart_timer.Enabled = true;//「再起動中」監視タイマー
			}

		}

		public static void SendText(int hWnd, string text)
		{
		    StringBuilder sb = new StringBuilder(text);
		    SendMessage((IntPtr)hWnd, WM_SETTEXT, (int)0, sb);
		}

        private void button1_Click(object sender, EventArgs e)
        {
			SequenceControl(1);
        }

        private void button2_Click(object sender, EventArgs e)
        {
			SequenceControl(2);
        }

        private void button3_Click(object sender, EventArgs e)
        {
			SequenceControl(3);
        }

        private void button4_Click(object sender, EventArgs e)
        {
			SequenceControl(4);
        }

        private void button5_Click(object sender, EventArgs e)
        {
			SequenceControl(5);
        }

        private void button6_Click(object sender, EventArgs e)
        {
			SequenceControl(6);
        }

        private void button7_Click(object sender, EventArgs e)
        {
			SequenceControl(7);
        }

        private void button8_Click(object sender, EventArgs e)
        {
			SequenceControl(8);
        }

        private void button9_Click(object sender, EventArgs e)
        {
			SequenceControl(9);
        }

        private void button10_Click(object sender, EventArgs e)
        {
			if(isExcuting)//実行中なら抜ける
			{
				return;
			}
			SequenceControl(0);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            string mes = "終了しますか？";
            DialogResult result = MessageBox.Show(mes, "リストアのアシスト", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (result == DialogResult.Yes)
            {
            }
            else
            {
                e.Cancel = true;
            }

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
			timer1.Interval = 3000;
            int path_len = path.Length;
            subFolders = System.IO.Directory.GetDirectories(path, "*", System.IO.SearchOption.TopDirectoryOnly);

            var next_list = new List<string>();

            for (int i = 0; i < subFolders.Length; i++)
            {
                string disp_name = subFolders[i].Substring(path_len + 1);
                next_list.Add(disp_name);
            }

			int prev_count = prev_list.Count();
			int next_count = next_list.Count();

			//projects配下にフォルダが何もないなら抜ける
            if(next_count == 0)
            {
                return;
            }

            if(prev_count == next_count)//以前と数が同じ
            {
                int count = 0;
                for (int i = 0; i < next_count; i++)
                {
					if(prev_list[i] == next_list[i])//フォルダ名が異なる箇所を数える
					{
						count++;
					}
				}

				if(count == next_count)//完全一致なら抜ける
				{
					return;
				}
				else
				{
	                prev_list.Clear();
	                for (int i = 0; i < next_list.Count; i++)
	                {
	                    prev_list.Add(next_list[i]);
	                }
				}
            }
            else
            {
                prev_list.Clear();
                for (int i = 0; i < next_list.Count; i++)
                {
                    prev_list.Add(next_list[i]);
                }
            }

            listView1.Items.Clear();
            for (int i = 0; i < next_list.Count; i++)
            {
				int num = i + 1;
				string number = num.ToString();
  				string[] item1 = {next_list[i], number};
              listView1.Items.Insert(0, new ListViewItem(item1));
            }
			listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
        }

		public void CheckRestartTitle(int hWnd)
		{
		    //ウィンドウタイトルを取得
		    StringBuilder title = new StringBuilder(256);
		    int titleLen = GetWindowText((int)hWnd, title, 256);
			label1.Text += title + "HIT" + "\r\n";

			//文字列の照合
            int hRestart = GetWindow((int)hWnd, GW_CHILD);
			do
			{
	            //ウィンドウのタイトルを取得する
	            StringBuilder tsb = new StringBuilder(256);
	            GetWindowText(hRestart, tsb, tsb.Capacity);
                string tsbStr = tsb.ToString();

                //ウィンドウのクラス名を取得する
                StringBuilder csb = new StringBuilder(256);
                GetClassName((IntPtr)hRestart, csb, csb.Capacity);
                string csbStr = csb.ToString();

				label1.Text += csbStr + "\r\n";
				label2.Text += tsbStr + "\r\n";

                if(tsbStr == "コントローラーを再起動しています")
                {
					isRestarting = true;
					label1.Text += tsbStr + "HIT" + "\r\n";
					break;
				}

                hRestart = GetWindow(hRestart, GW_HWNDNEXT);
		    }
		    while(hRestart != 0);
		}


        private void restart_timer_Tick(object sender, EventArgs e)
        {
			if(restart_timer_count > 30)//T.O
			{
				isRestarting = false;
				restart_timer_count = 0;
				restart_timer.Enabled = false;
				isExcuting = false;
				return;
			}

			if(!isRestarting)//画面が未表示の時
			{
	            Thread.Sleep(500);
	            IntPtr hWnd = FindWindow(null, "EPSON RC+ 7.0");
				if (hWnd != null)
				{
					CheckRestartTitle((int)hWnd);
				}
				else
				{
					label1.Text += "未表示　EPSON RC+ 7.0" + "\r\n";
				}
			}
			else//画面が表示中の時
			{
	            Thread.Sleep(500);
	            IntPtr hWnd = FindWindow(null, "EPSON RC+ 7.0");
				if (hWnd != null)
				{
					isRestarting = false;//一度初期化
					CheckRestartTitle((int)hWnd);
				}

				if(!isRestarting)
				{
					//「更新された実行用ファイルがあります」画面に対するOK押下
					Thread.Sleep(500);
					//Enterキー
					PostMessage((IntPtr)hWnd, WM_KEYDOWN, (int)VK_RETURN, 0);//Enterキーを押下
					Thread.Sleep(50);
					PostMessage((IntPtr)hWnd, WM_KEYUP, (int)VK_RETURN, 0);//Enterキーを戻す
					Thread.Sleep(50);

					//「プロジェクト同期」画面に対するOK「開始」押下
					Thread.Sleep(500);
					//Enterキー
		            IntPtr hprjsync = FindWindow(null, "プロジェクト同期");
					if (hprjsync != null)
					{
						PostMessage((IntPtr)hprjsync, WM_KEYDOWN, (int)VK_RETURN, 0);//Enterキーを押下
						Thread.Sleep(50);
						PostMessage((IntPtr)hprjsync, WM_KEYUP, (int)VK_RETURN, 0);//Enterキーを戻す
						Thread.Sleep(50);

						Thread.Sleep(5000);

						//「プロジェクト同期」画面に対する「閉じる」押下
						//Enterキー
						PostMessage((IntPtr)hprjsync, WM_KEYDOWN, (int)VK_RETURN, 0);//Enterキーを押下
						Thread.Sleep(50);
						PostMessage((IntPtr)hprjsync, WM_KEYUP, (int)VK_RETURN, 0);//Enterキーを戻す
						Thread.Sleep(50);
					}


					//終了
					label1.Text += "終了　EPSON RC+ 7.0" + "\r\n";
					isRestarting = false;
					restart_timer_count = 0;
					restart_timer.Enabled = false;
					term_timer.Enabled = true;
					return;
				}
			}

			restart_timer_count++;
        }

        private void term_timer_Tick(object sender, EventArgs e)
        {
			term_timer.Enabled = false;
			
            Thread.Sleep(500);
            IntPtr hWnd = FindWindow(null, "EPSON RC+ 7.0");
			if (hWnd != null)
			{
				//文字列の照合
		        int hRestart = GetWindow((int)hWnd, GW_CHILD);
				do
				{
		            //ウィンドウのタイトルを取得する
		            StringBuilder tsb = new StringBuilder(256);
		            GetWindowText(hRestart, tsb, tsb.Capacity);
		            string tsbStr = tsb.ToString();

		            //ウィンドウのクラス名を取得する
		            StringBuilder csb = new StringBuilder(256);
		            GetClassName((IntPtr)hRestart, csb, csb.Capacity);
		            string csbStr = csb.ToString();

					label1.Text += csbStr + "\r\n";
					label2.Text += tsbStr + "\r\n";

		            if(tsbStr == "システムリストアが完了しました。")
		            {
						label1.Text += tsbStr + "HIT" + "\r\n";
						break;
					}

		            hRestart = GetWindow(hRestart, GW_HWNDNEXT);
			    }
			    while(hRestart != 0);

	            //「システムリストアが完了しました」画面に対するOK押下
	            Thread.Sleep(200);
				//Enterキー
		        PostMessage((IntPtr)hWnd, WM_KEYDOWN, (int)VK_RETURN, 0);//Enterキーを押下
		        Thread.Sleep(50);
				PostMessage((IntPtr)hWnd, WM_KEYUP, (int)VK_RETURN, 0);//Enterキーを戻す
				Thread.Sleep(50);

	            //「メンテナンス」画面に対するESCキー押下で閉じる
	            Thread.Sleep(500);
				//Enterキー
		        PostMessage((IntPtr)holdHandle, WM_KEYDOWN, (int)VK_ESCAPE, 0);//Escキーを押下
		        Thread.Sleep(50);
				PostMessage((IntPtr)holdHandle, WM_KEYUP, (int)VK_ESCAPE, 0);//Escキーを戻す
				Thread.Sleep(50);
			}

			//main.prgの読み込み
            string [] prg_folders = System.IO.Directory.GetDirectories(subFolders[index], "*", System.IO.SearchOption.TopDirectoryOnly);
			string main_path = prg_folders[0] + "\\Main.prg";//元のファイル

            string[] strarray = File.ReadAllLines(main_path, Encoding.GetEncoding("SHIFT_JIS"));
            for (int i = 0; i < strarray.GetLength(0); i++)
            {
                if(strarray[i].Contains("For i"))
                {
		            string[] stringValues = strarray[i].Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
					numericUpDown1.Text = stringValues[3];//開始
					numericUpDown2.Text = stringValues[5];//終了
					break;
                }
            }

			isExcuting = false;
        }

        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
			//右クリックなら抜ける
		    if(e.Button == System.Windows.Forms.MouseButtons.Right)
			{
				return;
			}

			if(isExcuting)//実行中なら抜ける
			{
				return;
			}
			
			if(listView1.SelectedItems.Count > 0)//ListViewに1つでも登録がある
			{
	            label1.Text = "";
	            label2.Text = "";
			    index = listView1.SelectedItems[0].Index;//上から0オリジンで数えた位置
				isExcuting = true;

				//Epson RC+の制御
				SequenceControl(index + 1);
			}
        }

        private void listView1_MouseClick(object sender, MouseEventArgs e)
        {
			//右クリックなら抜ける
		    if(e.Button == System.Windows.Forms.MouseButtons.Right)
			{
				return;
			}

			if(isExcuting)//実行中なら抜ける
			{
				return;
			}
        }

        private void button11_Click(object sender, EventArgs e)
        {
			if(index == 0xFFFF)
			{
				string err = "最初に上の画面で型番を選択して下さい。";
                MessageBox.Show(err, "リストア　確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}

			if(numericUpDown1.Value >= numericUpDown2.Value)
			{
				string err = "移載範囲が不正です。確認して下さい。";
                MessageBox.Show(err, "移載範囲　確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}

			string mes = "移載範囲は" + numericUpDown1.Text + "から" + numericUpDown2.Text + "までに変更しますか？" + "\r\n\r\n" + 
			"変更するなら「はい」を押して下さい。ビルドします。" + "\r\n\r\n" + "間違えていたら「いいえ」を押して下さい。戻ります。";

            DialogResult result = MessageBox.Show(mes, "移載範囲　確認", MessageBoxButtons.YesNo);
			if(result == DialogResult.No)
			{
				return;
			}

			//Main.prgの書き換え
			string prg_path = "C:\\EpsonRC70\\projects\\T6_final";
			string main_path = prg_path + "\\Main.prg";//元のファイル
			string work_path = prg_path + "\\Main.tmp";//一時ファイル

            StringBuilder strread = new StringBuilder();
            string[] strarray = File.ReadAllLines(main_path, Encoding.GetEncoding("SHIFT_JIS"));
            for (int i = 0; i < strarray.GetLength(0); i++)
            {
                if(strarray[i].Contains("For i"))
                {
					strarray[i] = "	For i = " + numericUpDown1.Text + " To " + numericUpDown2.Text;
                }
				strread.AppendLine(strarray[i]);
            }
			File.WriteAllText(work_path, strread.ToString(), Encoding.GetEncoding("SHIFT_JIS"));

			//元のMain.prgファイルの存在確認
			if(System.IO.File.Exists(@main_path))
			{
				FileInfo file = new FileInfo(@main_path);
				if ((file.Attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)//読み取り専用なら解除
				{
					file.Attributes = FileAttributes.Normal;
				}

				System.IO.File.Delete(@main_path);//元のMain.prgファイルを削除
			}

            Thread.Sleep(200);//念の為

			File.Copy(@work_path, @main_path);//コピー


			//一時ファイルMain.tmpファイルの存在確認
			if(System.IO.File.Exists(@work_path))
			{
				FileInfo file = new FileInfo(@work_path);
				if ((file.Attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)//読み取り専用なら解除
				{
					file.Attributes = FileAttributes.Normal;
				}

				System.IO.File.Delete(@work_path);//一時Main.tmpファイルを削除
			}


			//ビルド開始
			//メニュー→プロジェクト→プロジェクトのリビルドをショートカットで遷移する。直接ファイルを直すとビルドが砂文字のままなのでリビルドにする
			//開いているMain.prgファイルを閉じないと、ビルドのショートカットが反映しない。タイミングか。
			int destWnd = 0;
            System.Diagnostics.Process[] process;
            process = System.Diagnostics.Process.GetProcessesByName("erc70");
            foreach (System.Diagnostics.Process ps in process)
            {
                destWnd = (int)ps.MainWindowHandle;//ターゲットのメインハンドル
                SetForegroundWindow(ps.MainWindowHandle);//ターゲットにFocusを当てる：非アクティブだとメニューにフォーカスが当たらない
            }

            PostMessage((IntPtr)destWnd, WM_SYSKEYDOWN, (int)VK_MENU, 0);//ALTキーを押下
            Thread.Sleep(200);
            PostMessage((IntPtr)destWnd, WM_SYSKEYUP, (int)VK_MENU, 0);//ALTキーを押下
            Thread.Sleep(200);

            PostMessage((IntPtr)destWnd, WM_KEYDOWN, (int)VK_P, 0);//Pキーを押下
            Thread.Sleep(200);
            PostMessage((IntPtr)destWnd, WM_KEYDOWN, (int)VK_R, 0);//Rキーを押下
            Thread.Sleep(200);

			//Main.prgが開いていたら閉じる
			//ファイル→ファイルを閉じる→をショートカットで遷移する
            PostMessage((IntPtr)destWnd, WM_SYSKEYDOWN, (int)VK_MENU, 0);//ALTキーを押下
            Thread.Sleep(200);
            PostMessage((IntPtr)destWnd, WM_SYSKEYUP, (int)VK_MENU, 0);//ALTキーを押下
            Thread.Sleep(200);

            PostMessage((IntPtr)destWnd, WM_KEYDOWN, (int)VK_F, 0);//Fキーを押下
            Thread.Sleep(200);
            PostMessage((IntPtr)destWnd, WM_KEYDOWN, (int)VK_C, 0);//Cキーを押下
            Thread.Sleep(200);

			//Main.prgが開いていないとメニューが開いたままの為、閉じるようにAltを一つ送信する
            PostMessage((IntPtr)destWnd, WM_SYSKEYDOWN, (int)VK_MENU, 0);//ALTキーを押下
            Thread.Sleep(200);
            PostMessage((IntPtr)destWnd, WM_SYSKEYUP, (int)VK_MENU, 0);//ALTキーを押下
            Thread.Sleep(200);

        }
    }
}
