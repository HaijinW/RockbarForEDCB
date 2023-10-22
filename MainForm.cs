using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Xml.XPath;
using EpgTimer;
using Nett;

namespace RockbarForEDCB
{
    /// <summary>
    /// メインフォームクラス
    /// </summary>
    public partial class MainForm : Form
    {
        // 設定情報
        private RockBarSetting rockbarSetting = null;

        // EpgTimerSrv接続可否
        private bool canConnect = true;

        // CtrlCmdの結果格納用
        private List<EpgServiceEventInfo> serviceEvents = new List<EpgServiceEventInfo>();
        private List<TunerReserveInfo> tunerReserveInfos = new List<TunerReserveInfo>();
        private List<ReserveData> reserveDatas = new List<ReserveData>();
        private List<RecFileInfo> recFileInfos = new List<RecFileInfo>();

        // サービス一覧(+番組)の保持用(TSID + SID → サービス情報(+番組))
        private Dictionary<string, EpgServiceEventInfo> serviceMap = new Dictionary<string, EpgServiceEventInfo>();

        // 番組一覧の保持用(TSID + SID + EventID → 番組情報)
        private Dictionary<string, EpgEventInfo> allEventMap = new Dictionary<string, EpgEventInfo>();

        // 予約情報の保持用(TSID + SID + EventID → 予約情報)
        private Dictionary<string, ReserveData> reserveMap = new Dictionary<string, ReserveData>();

        // 録画済み情報の保持用(TSID + SID + EventID → 録画済み情報)
        private Dictionary<uint, RecFileInfo> recMap = new Dictionary<uint, RecFileInfo>();

        // CSVサービスリストの格納
        private List<Service> allServiceList = null;
        private List<Service> favoriteServiceList = null;

        // CtrlCmdUtil
        private CtrlCmdUtil ctrlCmdUtil = new CtrlCmdUtil();

        // Rockbarから自動起動したTVTestのプロセス一覧
        private Dictionary<string, System.Diagnostics.Process> tvtestProcesses = new Dictionary<string, System.Diagnostics.Process>();

        // マウスのクリック位置を記憶
        private Point mousePoint;

        // 色の設定情報(設定情報から変数にロードしたもの)
        private Color formBackColor;
        private Color listBackColor;
        private Color okReserveListBackColor;
        private Color partialReserveListBackColor;
        private Color ngReserveListBackColor;
        private Color disabledReserveListBackColor;
        private Color listHeaderForeColor;
        private Color listHeaderBackColor;
        private Color foreColor;

        private Color menuBackColor;
        private Color okReserveMenuBackColor;
        private Color partialReserveMenuBackColor;
        private Color ngReserveMenuBackColor;
        private Color disabledReserveMenuBackColor;

        // 予約一覧のヘッダ数
        private int listHeaderCount = 0;

        // フィルタリング中かどうか
        private bool isFiltering => filteringLabel != null && filteringLabel.Visible;

        /// <summary>
        /// コンストラクタ
        /// コンフィグの読み込み・CtrlCmdの初期化・初回表示処理を行う。
        /// </summary>
        public MainForm()
        {
            InitializeComponent();

            try
            {
                rockbarSetting = Toml.ReadFile<RockBarSetting>(RockbarUtility.GetTomlSettingFilePath());
            }
            catch (FileNotFoundException)
            {
                // TOML設定ファイルが存在しない場合は準正常系として空設定で起動。それ以外の場合は例外を投げる
                rockbarSetting = new RockBarSetting();
            }

            // コンフィグファイルにwidth, height指定時のみ前回位置・サイズ・スプリッタ位置で起動
            if (rockbarSetting.Width != 0 && rockbarSetting.Height != 0)
            {
                this.StartPosition = FormStartPosition.Manual;
                this.Location = new Point(rockbarSetting.X, rockbarSetting.Y);
                this.Size = new Size(rockbarSetting.Width, rockbarSetting.Height);
                splitContainer.SplitterDistance = rockbarSetting.SplitterDistance;
            }

            // 設定反映
            applySetting();

            allServiceList = RockbarUtility.GetAllServicesFromSetting();
            favoriteServiceList = RockbarUtility.GetFavoriteServicesFromSetting();

            if (rockbarSetting.UseTcpIp) {
                // Pipe通信にする
                ctrlCmdUtil.SetSendMode(true);
                ctrlCmdUtil.SetNWSetting(rockbarSetting.IpAddress, rockbarSetting.PortNumber);
            }
            else
            {
                // Pipe通信にする
                ctrlCmdUtil.SetSendMode(false);
            }

            // 適当な通信を行って、通信可否を確認し問題があればメッセージを出す
            tunerReserveInfos.Clear();
            ErrCode errCode = ctrlCmdUtil.SendEnumTunerReserve(ref tunerReserveInfos);

            if (errCode != ErrCode.CMD_SUCCESS)
            {
                canConnect = false;
                MessageBox.Show(
                    $"EpgTimerSrvと接続できません。以降の通信を停止します。\nオプション設定を見直してアプリケーションを再起動してください。\n\nErrCode: {errCode}",
                    "EpgTimerSrv接続チェック失敗",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }

            // 初回表示
            RefreshEvent(true, true);

            // タイマーを有効化
            timer.Enabled = true;
        }

        /// <summary>
        /// ウィンドウプロシージャ
        /// リサイズ機能とダブルクリック無効化をwindowに追加
        /// 参照) https://stackoverflow.com/questions/31199437/borderless-and-resizable-form-c
        /// </summary>
        /// <param name="m">Windowsメッセージ</param>
        protected override void WndProc(ref Message m)
        {
            const int RESIZE_HANDLE_SIZE = 5;

            switch (m.Msg)
            {
                case 0x0084: /*NCHITTEST*/
                    base.WndProc(ref m);

                    if ((int) m.Result == 0x01) /*HTCLIENT*/
                    {
                        Point screenPoint = new Point(m.LParam.ToInt32());
                        Point clientPoint = this.PointToClient(screenPoint);
                        if (clientPoint.Y <= RESIZE_HANDLE_SIZE)
                        {
                            if (clientPoint.X <= RESIZE_HANDLE_SIZE)
                                m.Result = (IntPtr) 13; /*HTTOPLEFT*/
                            else if (clientPoint.X < (Size.Width - RESIZE_HANDLE_SIZE))
                                m.Result = (IntPtr) 12; /*HTTOP*/
                            else
                                m.Result = (IntPtr) 14; /*HTTOPRIGHT*/
                        }
                        else if (clientPoint.Y <= (Size.Height - RESIZE_HANDLE_SIZE))
                        {
                            if (clientPoint.X <= RESIZE_HANDLE_SIZE)
                                m.Result = (IntPtr) 10; /*HTLEFT*/
                            else if (clientPoint.X < (Size.Width - RESIZE_HANDLE_SIZE))
                                m.Result = (IntPtr) 2; /*HTCAPTION*/
                            else
                                m.Result = (IntPtr) 11; /*HTRIGHT*/
                        }
                        else
                        {
                            if (clientPoint.X <= RESIZE_HANDLE_SIZE)
                                m.Result = (IntPtr) 16; /*HTBOTTOMLEFT*/
                            else if (clientPoint.X < (Size.Width - RESIZE_HANDLE_SIZE))
                                m.Result = (IntPtr) 15; /*HTBOTTOM*/
                            else
                                m.Result = (IntPtr) 17; /*HTBOTTOMRIGHT*/
                        }
                    }
                    return;
                case 0x00A3: // WM_NCLBUTTONDBLCLK
                    // 非クライアント領域のダブルクリックによる最大化無効
                    m.Result = IntPtr.Zero;
                    return;
            }

            base.WndProc(ref m);
        }

        /// <summary>
        /// 設定反映処理
        /// 主に見た目部分の設定をフォームに反映する。初回起動時・設定変更時に実行
        /// </summary>
        private void applySetting()
        {
            // タスクトレイアイコン常時表示
            if (rockbarSetting.ShowTaskTrayIcon)
            {
                notifyIcon.Visible = true;
            }
            else
            {
                // 初回起動時・設定画面からの戻りで格納状態はないはず
                notifyIcon.Visible = false;
            }

            // 縦に並べて表示
            if (rockbarSetting.IsHorizontalSplit)
            {
                splitContainer.Orientation = Orientation.Horizontal;
            }
            else
            {
                splitContainer.Orientation = Orientation.Vertical;
            }

            // フォント
            TypeConverter fontConverter = TypeDescriptor.GetConverter(typeof(Font));
            Font font = (Font) fontConverter.ConvertFromString(rockbarSetting.Font);
            Font menuFont = (Font)fontConverter.ConvertFromString(rockbarSetting.MenuFont);
            Font tabFont = (Font)fontConverter.ConvertFromString(rockbarSetting.TabFont);
            Font buttonFont = (Font)fontConverter.ConvertFromString(rockbarSetting.ButtonFont);
            Font labelFont = (Font)fontConverter.ConvertFromString(rockbarSetting.LabelFont);
            Font textBoxFont = (Font)fontConverter.ConvertFromString(rockbarSetting.TextBoxFont);

            serviceListView.Font = font;
            tunerListView.Font = font;
            listContextMenuStrip.Font = menuFont;
            serviceTabControl.Font = tabFont;
            resetButton.Font = buttonFont;
            closeButton.Font = buttonFont;
            filteringLabel.Font = labelFont;
            filterTextBox.Font = textBoxFont;
            filterButton.Font = buttonFont;
            settingButton.Font = buttonFont;

            // フォント設定時にチャンネル一覧の2, 3列目の幅を設定(以降固定)
            // applySetting後に必ずServiceListはクリアされるため、1回2, 3列目だけのダミーデータを追加して列幅調整する
            String[] data = {
                "",
                "00:00-00:00",
                "◎",
                ""
            };

            ListViewItem item = new ListViewItem(data);
            // 調整時に目立つので背景色で隠す
            item.ForeColor = serviceListView.BackColor;
            serviceListView.Items.Add(item);
            // 2,3列目を内容で広げる
            serviceListView.AutoResizeColumn(1, ColumnHeaderAutoResizeStyle.ColumnContent);
            serviceListView.AutoResizeColumn(2, ColumnHeaderAutoResizeStyle.ColumnContent);

            adjustListViewColumns(serviceListView);
            adjustListViewColumns(tunerListView);

            // 色
            TypeConverter colorConverter = TypeDescriptor.GetConverter(typeof(Color));
            this.formBackColor = (Color)colorConverter.ConvertFromString(rockbarSetting.FormBackColor);
            this.listBackColor = (Color)colorConverter.ConvertFromString(rockbarSetting.ListBackColor);
            this.okReserveListBackColor = (Color)colorConverter.ConvertFromString(rockbarSetting.OkReserveListBackColor);
            this.partialReserveListBackColor = (Color)colorConverter.ConvertFromString(rockbarSetting.PartialReserveListBackColor);
            this.ngReserveListBackColor = (Color)colorConverter.ConvertFromString(rockbarSetting.NgReserveListBackColor);
            this.disabledReserveListBackColor = (Color)colorConverter.ConvertFromString(rockbarSetting.DisabledReserveListBackColor);
            this.listHeaderForeColor = (Color)colorConverter.ConvertFromString(rockbarSetting.ListHeaderForeColor);
            this.listHeaderBackColor = (Color)colorConverter.ConvertFromString(rockbarSetting.ListHeaderBackColor);
            this.foreColor = (Color)colorConverter.ConvertFromString(rockbarSetting.ForeColor);

            this.BackColor = this.formBackColor;

            this.menuBackColor = (Color)colorConverter.ConvertFromString(rockbarSetting.MenuBackColor);
            this.okReserveMenuBackColor = (Color)colorConverter.ConvertFromString(rockbarSetting.OkReserveMenuBackColor);
            this.partialReserveMenuBackColor = (Color)colorConverter.ConvertFromString(rockbarSetting.PartialReserveMenuBackColor);
            this.ngReserveMenuBackColor = (Color)colorConverter.ConvertFromString(rockbarSetting.NgReserveMenuBackColor);
            this.disabledReserveMenuBackColor = (Color)colorConverter.ConvertFromString(rockbarSetting.DisabledReserveMenuBackColor);

            serviceListView.BackColor = this.listBackColor;
            serviceListView.ForeColor = this.foreColor;

            tunerListView.BackColor = this.listBackColor;
            tunerListView.ForeColor = this.foreColor;

            listContextMenuStrip.BackColor = this.menuBackColor;
        }

        /// <summary>
        /// 描画更新処理
        /// 必要があればEpgTimerSrv通信を行い、チャンネルListView・チューナListViewの表示を更新する。
        /// </summary>
        /// <param name="isChannelRefresh">対象チャンネルリストの切り替え要否</param>
        /// <param name="isTrasnmission">EpgTimerSrvと通信要否</param>
        private void RefreshEvent(bool isChannelRefresh, bool isTrasnmission)
        {
            // フィルタリング中は何もしない
            if (!isChannelRefresh && isFiltering)
            {
                return;
            }

            // EpgTimerSrvと通信する
            if (isTrasnmission && canConnect)
            {
                // 予約一覧取得
                reserveDatas.Clear();
                ctrlCmdUtil.SendEnumReserve(ref reserveDatas);

                // 予約一覧関連のハッシュを作成
                reserveMap.Clear();
                foreach (ReserveData reserveData in reserveDatas)
                {
                    // TSID + SID + EventID → 予約情報
                    string tsidSidEventIdKey = RockbarUtility.GetKey(reserveData.TransportStreamID, reserveData.ServiceID, reserveData.EventID);

                    if (! reserveMap.ContainsKey(tsidSidEventIdKey))
                    {
                        reserveMap.Add(tsidSidEventIdKey, reserveData);
                    }
                }

                // 番組情報一覧取得
                serviceEvents.Clear();
                allEventMap.Clear();
                ctrlCmdUtil.SendEnumPgAll(ref serviceEvents);

                // 番組一覧関連のハッシュを作成
                serviceMap.Clear();
                foreach (EpgServiceEventInfo service in serviceEvents)
                {
                    // TSID + SID → サービス一覧(serviceが番組情報を保持している)
                    string tsidSidKey = RockbarUtility.GetKey(service.serviceInfo.TSID, service.serviceInfo.SID);

                    if (! serviceMap.ContainsKey(tsidSidKey))
                    {
                        serviceMap.Add(tsidSidKey, service);
                    }

                    // TSID + SID + EventID → 番組情報
                    foreach (EpgEventInfo ev in service.eventList)
                    {
                        string tsidSidEventIdKey = RockbarUtility.GetKey(ev.transport_stream_id, ev.service_id, ev.event_id);

                        if (! allEventMap.ContainsKey(tsidSidEventIdKey))
                        {
                            allEventMap.Add(tsidSidEventIdKey, ev);
                        }
                    }
                }

                // チューナーごとの予約一覧取得
                tunerReserveInfos.Clear();
                ctrlCmdUtil.SendEnumTunerReserve(ref tunerReserveInfos);

                // 録画済み情報の一覧取得
                if (serviceTabControl.SelectedTab == recTabPage)
                {
                    FetchRecList();
                }
                else
                {
                    recFileInfos.Clear();
                    recMap.Clear();
                }
            }

            // チャンネル一覧・チューナー一覧にキー情報だけでアイテムを表示する(対象放送波切り替え時も実行)
            if (isChannelRefresh)
            {
                if (serviceTabControl.SelectedTab == reserveTabPage)
                {
                    PrepareReserveList();
                }
                else if (serviceTabControl.SelectedTab == recTabPage)
                {
                    PrepareRecList();
                }
                else
                {
                    serviceListView.Items.Clear();

                    List<Service> services = null;

                    if (serviceTabControl.SelectedTab.Text == "お気に入り")
                    {
                        services = favoriteServiceList;
                    }
                    else
                    {
                        services = allServiceList;
                    }

                    // チャンネル表示
                    foreach (var service in services)
                    {
                        string key = RockbarUtility.GetKey(service.Tsid, service.Sid);

                        if (! serviceMap.ContainsKey(key))
                        {
                            // 定義したチャンネルがEDCBからのデータにない場合はスキップ(ユーザには設定画面でわかるようにする)
                            continue;
                        }

                        var matchedService = serviceMap[key];

                        if (serviceTabControl.SelectedTab == dttvTabPage && RockbarUtility.GetServiceType(matchedService.serviceInfo.ONID) != ServiceType.DTTV)
                        {
                            continue;
                        }
                        else if (serviceTabControl.SelectedTab == bsTabPage && RockbarUtility.GetServiceType(matchedService.serviceInfo.ONID) != ServiceType.BS)
                        {
                            continue;
                        }
                        else if (serviceTabControl.SelectedTab == csTabPage && RockbarUtility.GetServiceType(matchedService.serviceInfo.ONID) != ServiceType.CS)
                        {
                            continue;
                        }

                        // 1回キーとチャンネル名だけで追加
                        String[] data = {
                            matchedService.serviceInfo.service_name,
                            "",
                            "",
                            ""
                        };

                        ListViewItem item = new ListViewItem(data);
                        item.Name = key;

                        serviceListView.Items.Add(item);
                    }
                }

                adjustListViewColumns(serviceListView);

                filteringLabel.Visible = false;

                tunerListView.Items.Clear();

                // チューナー表示
                foreach (var tuner in tunerReserveInfos)
                {
                    string name = null;
                    
                    // 設定にチューナー名があれば取得し、なければデフォルト名で表示
                    if (rockbarSetting.BonDriverNameToTunerName.ContainsKey(tuner.tunerName))
                    {
                        name = rockbarSetting.BonDriverNameToTunerName[tuner.tunerName];
                    }
                    else
                    {
                        name = RockbarUtility.GetDefaultTunerName(tuner.tunerName);
                    }

                    // 連番付与
                    if (tuner.tunerID != 0xffffffff)
                    {
                        name += (tuner.tunerID & 0xffff).ToString();
                    }

                    // 1回キーとチューナー名だけで追加
                    string[] data = {name, ""};
                    ListViewItem item = new ListViewItem(data);
                    item.Name = tuner.tunerID.ToString();

                    tunerListView.Items.Add(item);
                }

                adjustListViewColumns(tunerListView);
            }

            if (serviceTabControl.SelectedTab == reserveTabPage)
            {
                RefreshReserveList();
            }
            else if (serviceTabControl.SelectedTab == recTabPage)
            {
                RefreshRecList();
            }
            else
            {
                // チャンネル一覧の追加済みアイテムに対して、現在放送中の番組情報を付与する
                foreach (var service in serviceEvents)
                {
                    string key = RockbarUtility.GetKey(service.serviceInfo.TSID, service.serviceInfo.SID);

                    var ev = service.eventList.Find(x => x.start_time <= DateTime.Now && x.start_time.AddSeconds(x.durationSec) >= DateTime.Now);

                    // チャンネル一覧ListViewでこのサービスが表示対象
                    if (serviceListView.Items.ContainsKey(key))
                    {
                        ListViewItem item = serviceListView.Items[key];

                        if (ev != null)
                        {
                            // 現在放送中の番組あり
                            string eventKey = RockbarUtility.GetKey(ev.transport_stream_id, ev.service_id, ev.event_id);

                            ReserveStatus reserveStatus = ReserveStatus.NONE;

                            // 予約状態文字列を取得
                            if (reserveMap.ContainsKey(eventKey)) {
                                ReserveData reserveData = reserveMap[eventKey];

                                reserveStatus = ReserveStatus.OK;

                                if (reserveData.RecSetting.IsNoRec())
                                {
                                    reserveStatus = ReserveStatus.DISABLED;
                                }
                                else if (reserveData.OverlapMode == 1)
                                {
                                    reserveStatus = ReserveStatus.PARTIAL;
                                }
                                else if (reserveData.OverlapMode == 2)
                                {
                                    reserveStatus = ReserveStatus.NG;
                                }
                            }

                            string reserveString = RockbarUtility.GetReserveStatusString(reserveStatus);

                            // 番組情報を表示
                            item.SubItems[1].Text = ev.start_time.ToString("HH:mm") + '-' + ev.start_time.AddSeconds(ev.durationSec).ToString("HH:mm");
                            item.SubItems[2].Text = reserveString;
                            item.SubItems[3].Text = ev.ShortInfo?.event_name;
                            item.ToolTipText = ev.ShortInfo?.event_name;

                            // 色変更
                            switch (reserveStatus)
                            {
                                case ReserveStatus.NONE:
                                    item.BackColor = this.listBackColor;
                                    break;
                                case ReserveStatus.OK:
                                    item.BackColor = this.okReserveListBackColor;
                                    break;
                                case ReserveStatus.PARTIAL:
                                    item.BackColor = this.partialReserveListBackColor;
                                    break;
                                case ReserveStatus.NG:
                                    item.BackColor = this.ngReserveListBackColor;
                                    break;
                                case ReserveStatus.DISABLED:
                                    item.BackColor = this.disabledReserveListBackColor;
                                    break;
                            }
                        }
                        else
                        {
                            // 現在放送中の番組なし
                            item.SubItems[1].Text = "";
                            item.SubItems[2].Text = "";
                            item.SubItems[3].Text = "";
                            item.ToolTipText = "";
                            item.BackColor = this.listBackColor;
                        }
                    }
                }
            }

            // チューナー一覧の追加済みアイテムに対して、現在録画中の番組情報を付与する
            foreach (var tuner in tunerReserveInfos)
            {
                // チューナーのListViewItemを拾う
                var item = tunerListView.Items[tuner.tunerID.ToString()];

                // 使用中の予約タイトルとツールチップを表示
                HashSet<uint> reserveIds = tuner.reserveList.ToHashSet();

                List<ReserveData> reserves = reserveDatas.FindAll(x => reserveIds.Contains(x.ReserveID));

                var currentReserve = reserves.Find(x => x.StartTime <= DateTime.Now && x.StartTime.AddSeconds(x.DurationSecond) >= DateTime.Now);

                if (currentReserve != null)
                {
                    item.SubItems[1].Text = currentReserve.Title;
                    item.ToolTipText = currentReserve.Title; ;
                }
                else
                {
                    item.SubItems[1].Text = "";
                    item.ToolTipText = "";
                }

                // 直近30件に限らず、将来変な予約がある場合警告として色を変える
                if (reserves.Count(x => x.OverlapMode == 1) > 0)
                {
                    // 一部予約に1件でも予約が入っている場合、黃背景色で警告
                    item.BackColor = this.partialReserveListBackColor;
                }
                else if (reserves.Count(x => x.OverlapMode == 2) > 0)
                {
                    // TU不足に1件でも予約が入っている場合、赤背景色で警告
                    item.BackColor = this.ngReserveListBackColor;
                }
                else
                {
                    item.BackColor = this.listBackColor;
                }
            }
        }

        /// <summary>
        /// 予約一覧の描画準備処理
        /// </summary>
        private void PrepareReserveList()
        {
            serviceListView.Items.Clear();
            listHeaderCount = 0;

            // 一覧表示
            string dateTextCache = "";
            foreach (var reserveData in reserveDatas.OrderBy(r => r.StartTime))
            {
                DateTime startTime = reserveData.StartTime;
                DateTime endTime = startTime.AddSeconds(reserveData.DurationSecond);
                if (endTime < DateTime.Now)
                {
                    continue;
                }

                string onAirDate = startTime.ToString("yyyy/MM/dd(ddd)");
                if (onAirDate != dateTextCache)
                {
                    String[] dateData = {
                        onAirDate,
                        "",
                        "",
                        ""
                    };

                    ListViewItem dateItem = new ListViewItem(dateData);
                    dateItem.ForeColor = this.listHeaderForeColor;
                    dateItem.BackColor = this.listHeaderBackColor;

                    serviceListView.Items.Add(dateItem);
                    dateTextCache = onAirDate;
                    listHeaderCount++;
                }

                String[] data = {
                    reserveData.StationName,
                    "",
                    "",
                    ""
                };

                ListViewItem item = new ListViewItem(data);
                serviceListView.Items.Add(item);
            }

            adjustListViewColumns(serviceListView);
        }

        /// <summary>
        /// 予約一覧の描画更新処理
        /// </summary>
        private void RefreshReserveList()
        {
            if (serviceListView.Items.Count < reserveDatas.Count + listHeaderCount)
            {
                PrepareReserveList();
            }

            // 一覧表示
            int index = 0;
            int listViewCount = serviceListView.Items.Count;
            string dateTextCache = "";
            foreach (var reserveData in reserveDatas.OrderBy(r => r.StartTime))
            {
                DateTime startTime = reserveData.StartTime;
                DateTime endTime = startTime.AddSeconds(reserveData.DurationSecond);
                if (endTime < DateTime.Now)
                {
                    continue;
                }

                string onAirDate = startTime.ToString("yyyy/MM/dd(ddd)");
                if (onAirDate != dateTextCache)
                {
                    ListViewItem dateItem = serviceListView.Items[index];
                    dateItem.SubItems[0].Text = onAirDate;
                    dateItem.SubItems[1].Text = "時間";
                    dateItem.SubItems[3].Text = "番組名";
                    dateItem.ForeColor = this.listHeaderForeColor;
                    dateItem.BackColor = this.listHeaderBackColor;

                    dateTextCache = onAirDate;
                    index++;
                    if (index >= listViewCount)
                    {
                        break;
                    }
                }

                ReserveStatus reserveStatus = ReserveStatus.OK;

                if (reserveData.RecSetting.IsNoRec())
                {
                    reserveStatus = ReserveStatus.DISABLED;
                }
                else if (reserveData.OverlapMode == 1)
                {
                    reserveStatus = ReserveStatus.PARTIAL;
                }
                else if (reserveData.OverlapMode == 2)
                {
                    reserveStatus = ReserveStatus.NG;
                }

                ListViewItem item = serviceListView.Items[index];
                item.SubItems[0].Text = reserveData.StationName;
                item.SubItems[1].Text = startTime.ToString("HH:mm") + '-' + endTime.ToString("HH:mm");
                item.SubItems[2].Text = RockbarUtility.GetReserveStatusString(reserveStatus);
                item.SubItems[3].Text = reserveData.Title;
                item.Name = RockbarUtility.GetKey(reserveData.TransportStreamID, reserveData.ServiceID, reserveData.EventID);
                item.ToolTipText = $"{startTime.ToString("yyyy/MM/dd(ddd) HH:mm")}-{endTime.ToString("HH:mm")} {reserveData.Title}";
                item.ForeColor = this.foreColor;

                // 無効予約の場合
                if (reserveStatus == ReserveStatus.DISABLED)
                {
                    item.BackColor = this.disabledReserveListBackColor;
                }
                // 変な予約がある場合警告として色を変える
                else if (reserveStatus == ReserveStatus.PARTIAL)
                {
                    // 一部予約の場合、黃背景色で警告
                    item.BackColor = this.partialReserveListBackColor;
                }
                else if (reserveStatus == ReserveStatus.NG)
                {
                    // TU不足の場合、赤背景色で警告
                    item.BackColor = this.ngReserveListBackColor;
                }
                else if (startTime <= DateTime.Now && endTime >= DateTime.Now)
                {
                    // 現在録画中の場合、正常予約背景色で表示
                    item.BackColor = this.okReserveListBackColor;
                }
                else
                {
                    item.BackColor = this.listBackColor;
                }

                index++;
                if (index >= listViewCount)
                {
                    break;
                }
            }

            while (index < serviceListView.Items.Count)
            {
                serviceListView.Items.RemoveAt(index);
            }
        }

        /// <summary>
        /// 録画済み一覧の取得処理
        /// </summary>
        private void FetchRecList()
        {
            recFileInfos.Clear();
            recMap.Clear();
            ctrlCmdUtil.SendEnumRecInfoBasic(ref recFileInfos);

            int recListMaxCount = rockbarSetting.RecListMaxCount;
            if (recListMaxCount > 0 && recFileInfos.Count > recListMaxCount)
            {
                recFileInfos = recFileInfos.GetRange(recFileInfos.Count - recListMaxCount, recListMaxCount);
            }
            recFileInfos.Reverse();

            recMap = recFileInfos.ToDictionary(r => r.ID);
        }

        /// <summary>
        /// 録画済み一覧の描画準備処理
        /// </summary>
        private void PrepareRecList()
        {
            if (recFileInfos.Count == 0)
            {
                FetchRecList();
            }
            
            serviceListView.Items.Clear();
            listHeaderCount = 0;

            // 一覧表示
            string dateTextCache = "";
            foreach (var recFile in recFileInfos)
            {
                DateTime startTime = recFile.StartTime;
                string onAirDate = startTime.ToString("yyyy/MM/dd(ddd)");
                if (onAirDate != dateTextCache)
                {
                    String[] dateData = {
                        onAirDate,
                        "",
                        "",
                        ""
                    };

                    ListViewItem dateItem = new ListViewItem(dateData);
                    dateItem.ForeColor = this.listHeaderForeColor;
                    dateItem.BackColor = this.listHeaderBackColor;

                    serviceListView.Items.Add(dateItem);
                    dateTextCache = onAirDate;
                    listHeaderCount++;
                }

                String[] data = {
                    recFile.ServiceName,
                    "",
                    "",
                    ""
                };

                ListViewItem item = new ListViewItem(data);
                serviceListView.Items.Add(item);
            }

            adjustListViewColumns(serviceListView);
        }

        /// <summary>
        /// 録画済み一覧の描画更新処理
        /// </summary>
        private void RefreshRecList()
        {
            if (serviceListView.Items.Count < recFileInfos.Count + listHeaderCount)
            {
                PrepareRecList();
            }

            // 一覧表示
            int index = 0;
            int listViewCount = serviceListView.Items.Count;
            string dateTextCache = "";
            foreach (var recFile in recFileInfos)
            {
                DateTime startTime = recFile.StartTime;
                DateTime endTime = startTime.AddSeconds(recFile.DurationSecond);

                string onAirDate = startTime.ToString("yyyy/MM/dd(ddd)");
                if (onAirDate != dateTextCache)
                {
                    ListViewItem dateItem = serviceListView.Items[index];
                    dateItem.SubItems[0].Text = onAirDate;
                    dateItem.SubItems[1].Text = "時間";
                    dateItem.SubItems[3].Text = "番組名";
                    dateItem.ForeColor = this.listHeaderForeColor;
                    dateItem.BackColor = this.listHeaderBackColor;

                    dateTextCache = onAirDate;
                    index++;
                    if (index >= listViewCount)
                    {
                        break;
                    }
                }

                RecEndStatus recEndStatus = (RecEndStatus) recFile.RecStatus;

                ListViewItem item = serviceListView.Items[index];
                item.SubItems[0].Text = recFile.ServiceName;
                item.SubItems[1].Text = startTime.ToString("HH:mm") + '-' + endTime.ToString("HH:mm");
                item.SubItems[2].Text = RockbarUtility.GetRecEndStatusString(recEndStatus);
                item.SubItems[3].Text = recFile.Title;
                item.Name = recFile.ID.ToString();
                item.ToolTipText = createRecInfoTooltipTexts(recFile);
                item.ForeColor = this.foreColor;

                // 録画結果が正常の場合
                if (recEndStatus == RecEndStatus.NORMAL || recEndStatus == RecEndStatus.CHG_TIME || recEndStatus == RecEndStatus.NEXT_START_END)
                {
                    if (recFile.Scrambles > 0)
                    {
                        // スクランブル解除漏れありの場合、黃背景色で警告
                        item.BackColor = this.partialReserveListBackColor;
                    }
                    else if (recFile.Drops > 0)
                    {
                        // ドロップありの場合、赤背景色で警告
                        item.BackColor = this.ngReserveListBackColor;
                    }
                    else
                    {
                        item.BackColor = this.listBackColor;
                    }
                }
                // 録画結果が正常ではない場合
                else if (recEndStatus == RecEndStatus.END_SUBREC)
                {
                    // サブフォルダへの録画の場合、ダークスレートグレー背景色で通知
                    item.BackColor = this.okReserveListBackColor;
                }
                else if (recEndStatus == RecEndStatus.ERR_END || recEndStatus == RecEndStatus.END_SUBREC || recEndStatus == RecEndStatus.NOT_START_HEAD)
                {
                    // 録画中のエラー、一部のみ録画の場合、黃背景色で警告
                    item.BackColor = this.partialReserveListBackColor;
                }
                else if (recEndStatus == RecEndStatus.NO_RECMODE)
                {
                    // 無効扱いの場合、グレー背景色で通知
                    item.BackColor = this.disabledReserveListBackColor;
                }
                else
                {
                    // それ以外のエラーの場合、赤背景色で警告
                    item.BackColor = this.ngReserveListBackColor;
                }

                index++;
                if (index >= listViewCount)
                {
                    break;
                }
            }

            while (index < serviceListView.Items.Count)
            {
                serviceListView.Items.RemoveAt(index);
            }
        }

        /// <summary>
        /// 設定ファイルの再読み込み処理
        /// </summary>
        private void ReloadSetting()
        {
            // 設定ファイルを書き込んだあとの読み込みなので基本的に例外は発生しないはず。発生した場合は例外を投げる
            rockbarSetting = Toml.ReadFile<RockBarSetting>(RockbarUtility.GetTomlSettingFilePath());

            allServiceList = RockbarUtility.GetAllServicesFromSetting();
            favoriteServiceList = RockbarUtility.GetFavoriteServicesFromSetting();
        }

        /// <summary>
        /// フィルタ処理
        /// チャンネル一覧から、フィルタ文字列にチャンネル名も番組名も一致しないチャンネルを削除して表示更新する(大文字・小文字無視)
        /// </summary>
        private void Filter()
        {
            if (string.IsNullOrEmpty(filterTextBox.Text))
            {
                ResetFilter();
                return;
            }

            for (var i = serviceListView.Items.Count - 1; i >= 0; i--)
            {
                var item = serviceListView.Items[i];

                if (string.IsNullOrEmpty(item.Name))
                {
                    var lastIndex = serviceListView.Items.Count - 1;
                    if (i < lastIndex)
                    {
                        var prevItem = serviceListView.Items[i + 1];
                        if (string.IsNullOrEmpty(prevItem.Name))
                        {
                            serviceListView.Items.RemoveAt(i);
                        }
                    }
                    else if (i == lastIndex)
                    {
                        serviceListView.Items.RemoveAt(i);
                    }
                    continue;
                }
                else if (
                    item.SubItems[0].Text.IndexOf(filterTextBox.Text, StringComparison.OrdinalIgnoreCase) < 0 &&
                    item.SubItems[3].Text.IndexOf(filterTextBox.Text, StringComparison.OrdinalIgnoreCase) < 0
                )
                {
                    serviceListView.Items.RemoveAt(i);
                }
            }

            filteringLabel.Visible = true;

            serviceListView.Refresh();
        }

        /// <summary>
        /// フィルタのリセット処理
        /// フィルタ文字列をクリアしフィルタ結果をリセットする。
        /// </summary>
        private void ResetFilter()
        {
            filterTextBox.Clear();
            filteringLabel.Visible = false;
            RefreshEvent(true, false);
        }

        /// <summary>
        /// 録画モード有効・無効の切り替え処理
        /// 録画有効時は無効化、無効時は有効化する。
        /// </summary>
        private void ToggleRecMode(ReserveData reserve)
        {
            if (reserve.RecSetting.IsNoRec())
            {
                reserve.RecSetting.RecMode = rockbarSetting.FixNoRecToServiceOnly ? (byte)1 : reserve.RecSetting.GetRecMode();
            }
            else
            {
                // 録画モード情報を維持して無効化
                var recMode = reserve.RecSetting.RecMode;
                reserve.RecSetting.RecMode = (byte)(rockbarSetting.FixNoRecToServiceOnly ? 5 : 5 + (recMode + 4) % 5);
            }

            var err = ctrlCmdUtil.SendChgReserve(new List<ReserveData>() { reserve });
            if (err != ErrCode.CMD_SUCCESS)
            {
                MessageBox.Show("予約変更でエラーが発生しました。", "予約変更エラー");
            }
            RefreshEvent(false, true);
        }

        /// <summary>
        /// 番組の右クリックコンテキストメニューItem作成処理
        /// 番組情報・予約情報からチャンネル一覧・チューナー一覧用のコンテキストメニューitemを作成する。
        /// </summary>
        /// <param name="ev">番組情報</param>
        /// <param name="reserve">予約情報</param>
        /// <param name="isTuner">チューナー一覧用？</param>
        /// <returns></returns>
        private ToolStripMenuItem createEventToolStripMenuItem(EpgEventInfo ev, ReserveData reserve, bool isTuner)
        {
            ReserveStatus reserveStatus = ReserveStatus.NONE;

            // 予約ステータスを判別
            if (reserve != null)
            {
                if (reserve.RecSetting.IsNoRec())
                {
                    reserveStatus = ReserveStatus.DISABLED;
                }
                else if (reserve.OverlapMode == 0)
                {
                    if (ev == null)
                    {
                        reserveStatus = ReserveStatus.DISAPPEARED;
                    }
                    else
                    {
                        reserveStatus = ReserveStatus.OK;
                    }
                }
                else if (reserve.OverlapMode == 1)
                {
                    reserveStatus = ReserveStatus.PARTIAL;
                }
                else
                {
                    reserveStatus = ReserveStatus.NG;
                }
            }

            ToolStripMenuItem item = new ToolStripMenuItem();

            // 録画ステータスに異常があれば色を変える
            switch (reserveStatus)
            {
                case ReserveStatus.OK:
                    if (!isTuner)
                    {
                        item.BackColor = this.okReserveMenuBackColor;
                    }
                    break;
                case ReserveStatus.PARTIAL:
                    item.BackColor = this.partialReserveMenuBackColor;
                    break;
                case ReserveStatus.NG:
                    item.BackColor = this.ngReserveMenuBackColor;
                    break;
                case ReserveStatus.DISABLED:
                    item.BackColor = this.disabledReserveMenuBackColor;
                    break;
            }

            string reserveString = RockbarUtility.GetReserveStatusString(reserveStatus);

            if (isTuner)
            {
                // チューナーから開いた場合は予約情報を表示(必ず予約情報あり)
                item.Text = reserve.StartTime.ToString("MM/dd HH:mm") + "～" + reserve.StartTime.AddSeconds(reserve.DurationSecond).ToString("HH:mm") + "  " +
                    reserveString + "    " + reserve.StationName + "    " + reserve.Title;
            }
            else
            {
                // サービスから開いた場合は番組情報を表示(必ず番組情報あり)
                item.Text = ev.start_time.ToString("MM/dd HH:mm") + "～" + ev.start_time.AddSeconds(ev.durationSec).ToString("HH:mm") + "  " +
                    reserveString + "  " + ev.ShortInfo?.event_name;
            }

            // 番組情報がある場合はサブメニューに番組情報を追加
            if (ev != null)
            {
                // Webリンク使用時のみ1行目にWebリンク用のボタンを表示
                if (rockbarSetting.UseWebLink)
                {
                    item.DropDownItems.Add(">>");
                    item.DropDownItems[item.DropDownItems.Count - 1].Click += (s2, e2) => accessWebUrl(ev);

                    item.DropDownItems.Add(new ToolStripSeparator());
                }

                // 基本情報をサブメニューに追加
                var shortStrs = RockbarUtility.BreakString(ev.ShortInfo?.text_char);

                if (shortStrs != null)
                {
                    foreach (string str in shortStrs)
                    {
                        item.DropDownItems.Add(str);
                        item.DropDownItems[item.DropDownItems.Count - 1].Enabled = false;
                    }
                }

                item.DropDownItems.Add(new ToolStripSeparator());

                // 拡張情報をサブメニューに追加
                var longStrs = RockbarUtility.BreakString(ev.ExtInfo?.text_char);

                if (longStrs != null)
                {
                    foreach (string str in longStrs)
                    {
                        item.DropDownItems.Add(str);
                        item.DropDownItems[item.DropDownItems.Count - 1].Enabled = false;
                    }
                }
            }

            return item;
        }

        /// <summary>
        /// 録画済み情報のテキスト生成処理
        /// </summary>
        /// <param name="recFile">録画済み情報</param>
        private List<List<string>> createRecInfoDetailTexts(RecFileInfo recFile)
        {
            List<List<string>> detailTexts = new List<List<string>>();
            var dateTimes = new List<string>() { $"{recFile.StartTime.ToString("yyyy/MM/dd(ddd) HH:mm")}-{recFile.StartTime.AddSeconds(recFile.DurationSecond).ToString("HH:mm")}" };
            detailTexts.Add(dateTimes);

            var services = new List<string>() { recFile.ServiceName, recFile.Title };
            detailTexts.Add(services);

            var results = new List<string>() { $"結果 : {recFile.Comment}" };
            results.AddRange(RockbarUtility.BreakString($"録画ファイル : {recFile.RecFilePath}", 50));
            detailTexts.Add(results);

            var ids = new List<string>() {
                $"OriginalNetworkID : {recFile.OriginalNetworkID} (0x{recFile.OriginalNetworkID.ToString("X4")})",
                $"TransportStreamID : {recFile.TransportStreamID} (0x{recFile.TransportStreamID.ToString("X4")})",
                $"ServiceID : {recFile.ServiceID} (0x{recFile.ServiceID.ToString("X4")})",
                $"EventID : {recFile.EventID} (0x{recFile.EventID.ToString("X4")})"
            };
            detailTexts.Add(ids);

            var scrambles = new List<string>() {
                $"Drop : {recFile.Drops}",
                $"Scramble : {recFile.Scrambles}"
            };
            detailTexts.Add(scrambles);

            return detailTexts;
        }

        /// <summary>
        /// 録画済み情報のツールチップテキスト生成処理
        /// </summary>
        /// <param name="recFile">録画済み情報</param>
        private string createRecInfoTooltipTexts(RecFileInfo recFile)
        {
            var detailTexts = createRecInfoDetailTexts(recFile);
            return string.Join("\n\n", detailTexts.Select(t => string.Join("\n", t)));
        }

        /// <summary>
        /// Web番組詳細にアクセスする
        /// </summary>
        /// <param name="ev">番組情報</param>
        private void accessWebUrl(EpgEventInfo ev)
        {
            string url = rockbarSetting.WebLinkUrl;
            url = url.Replace("{ONID}", ev.original_network_id.ToString());
            url = url.Replace("{TSID}", ev.transport_stream_id.ToString());
            url = url.Replace("{SID}", ev.service_id.ToString());
            url = url.Replace("{EID}", ev.event_id.ToString());

            try
            {
                // URLがパースできるかどうかを事前判定しておく
                new Uri(url);

                var startInfo = new System.Diagnostics.ProcessStartInfo(url);
                startInfo.UseShellExecute = true;
                System.Diagnostics.Process.Start(startInfo);
            }
            catch
            {
                MessageBox.Show($"Web番組詳細URLが不正です。Web番組詳細URLの設定を見直してください。\nURL: {url}", "ブラウザ起動エラー");
                return;
            }
        }

        /// <summary>
        /// Web番組詳細にアクセスする
        /// </summary>
        /// <param name="recFile">録画済み情報</param>
        private void accessWebUrl(RecFileInfo recFile)
        {
            string url = rockbarSetting.RecInfoWebLinkUrl;
            url = url.Replace("{RecID}", recFile.ID.ToString());

            try
            {
                // URLがパースできるかどうかを事前判定しておく
                new Uri(url);

                var startInfo = new System.Diagnostics.ProcessStartInfo(url);
                startInfo.UseShellExecute = true;
                System.Diagnostics.Process.Start(startInfo);
            }
            catch
            {
                MessageBox.Show($"Web番組詳細URLが不正です。Web番組詳細URLの設定を見直してください。\nURL: {url}", "ブラウザ起動エラー");
                return;
            }
        }

        /// <summary>
        /// リストビューカラム幅調整処理
        /// 1列目を内容に合わせ、最終列コントロールいっぱいまで広げる。
        /// </summary>
        /// <param name="listView">対象リストビュー</param>
        private void adjustListViewColumns(ListView listView)
        {
            // 1列目を内容で広げる
            listView.AutoResizeColumn(0, ColumnHeaderAutoResizeStyle.ColumnContent);

            // 1列目・最終列以外は固定幅
            int columnWidth = 0;

            for (var i = 0; i < listView.Columns.Count - 1; i++)
            {
                columnWidth += listView.Columns[i].Width;
            }

            // 最終列をいっぱいに広げる
            listView.Columns[listView.Columns.Count - 1].Width = listView.ClientRectangle.Width - columnWidth - 2;
        }

        /// <summary>
        /// TVTest起動処理
        /// TSID, SIDを指定してTVTestを起動する。地デジ・BS/CSで異なるオプションを使用する。
        /// </summary>
        /// <param name="isDttv">地デジ？</param>
        /// <param name="tsid">TSID</param>
        /// <param name="sid">SID</param>
        /// <returns>TVTestプロセス</returns>
        private System.Diagnostics.Process startTvTest(bool isDttv, uint tsid, uint sid)
        {
            System.Diagnostics.Process result = null;

            try
            {
                if (isDttv)
                {
                    result = System.Diagnostics.Process.Start(rockbarSetting.TvtestPath, $"{rockbarSetting.TvtestDttvOption} /tsid {tsid} /sid {sid}");
                }
                else
                {
                    result = System.Diagnostics.Process.Start(rockbarSetting.TvtestPath, $"{rockbarSetting.TvtestBscsOption} /tsid {tsid} /sid {sid}");
                }
            }
            catch
            {
                MessageBox.Show("TVTestの起動に失敗しました。TVTestの設定を見直してください。", "TVTest起動エラー");
            }

            return result;
        }

        /// <summary>
        /// TvtPlayプラグインを有効化してTVTest起動処理
        /// ファイルパスとTvtPlayの起動オプションを指定してTVTestを起動する。
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        /// <returns>TVTestプロセス</returns>
        private System.Diagnostics.Process startTvtPlay(string filePath)
        {
            System.Diagnostics.Process result = null;

            try
            {
                result = System.Diagnostics.Process.Start(rockbarSetting.TvtestPath, $"{rockbarSetting.TvtestTsFileOption} \"{filePath}\"");
            }
            catch
            {
                MessageBox.Show("TVTestの起動に失敗しました。TVTestの設定を見直してください。", "TVTest起動エラー");
            }

            return result;
        }

        /// <summary>
        /// チャンネル一覧マウスクリック時処理
        /// 右クリック時、直近30件の番組情報をコンテキストメニューに表示。
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void serviceListView_MouseClick(object sender, MouseEventArgs e)
        {
            if (serviceTabControl.SelectedTab == reserveTabPage)
            {
                reserveListView_MouseClick(sender, e);
                return;
            }
            else if (serviceTabControl.SelectedTab == recTabPage)
            {
                recListView_MouseClick(sender, e);
                return;
            }

            // 右クリック
            // 現在の番組を含め、今後の番組を30件までコンテキストメニューで表示(TVRockの仕様踏襲)
            if (e.Button == MouseButtons.Right)
            {
                // クリックした箇所が自動選択されるので拾う
                var selected = serviceListView.SelectedItems[0];

                listContextMenuStrip.Items.Clear();

                EpgServiceEventInfo sv = serviceMap[selected.Name];

                var afterEventList = sv.eventList.FindAll(x => x.start_time.AddSeconds(x.durationSec) >= DateTime.Now).OrderBy(a => a.start_time);

                int i = 0;

                foreach (var ev in afterEventList)
                {
                    string eventKey = RockbarUtility.GetKey(ev.transport_stream_id, ev.service_id, ev.event_id);

                    ReserveData reserveData = null;
                    
                    if (reserveMap.ContainsKey(eventKey))
                    {
                        reserveData = reserveMap[eventKey];
                    }

                    listContextMenuStrip.Items.Add(createEventToolStripMenuItem(ev, reserveData, false));

                    i++;

                    if (i >= 30)
                    {
                        break;
                    }
                }

                listContextMenuStrip.Show((Control) sender, new Point(0, e.Y) );
            }
        }

        /// <summary>
        /// チャンネル一覧マウスアップ時処理
        /// 中央ボタンのクリック検出用。
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void serviceListView_MouseUp(object sender, MouseEventArgs e)
        {
            if (serviceTabControl.SelectedTab == reserveTabPage)
            {
                reserveListView_MouseUp(sender, e);
                return;
            }
        }

        /// <summary>
        /// チャンネル一覧マウスダブルクリック処理
        /// TVTest使用オプションがONの場合、カーソル箇所の番組を対象にTVTestを起動する。
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void serviceListView_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (serviceTabControl.SelectedTab == reserveTabPage)
            {
                reserveListView_MouseDoubleClick(sender, e);
                return;
            }
            else if (serviceTabControl.SelectedTab == recTabPage)
            {
                recListView_MouseDoubleClick(sender, e);
                return;
            }

            // TVTest使用時のみ
            if (! rockbarSetting.UseDoubleClickTvtest)
            {
                return;
            }

            // 左ダブルクリック
            // TVTestを起動する
            if (e.Button == MouseButtons.Left)
            {
                // クリックした箇所が自動選択されるので拾う
                var selected = serviceListView.SelectedItems[0];

                EpgServiceEventInfo sv = serviceMap[selected.Name];

                if (RockbarUtility.GetServiceType(sv.serviceInfo.ONID) == ServiceType.DTTV) {
                    // 地上波
                    startTvTest(true, sv.serviceInfo.TSID, sv.serviceInfo.SID);
                }
                else
                {
                    // BS, CS
                    startTvTest(false, sv.serviceInfo.TSID, sv.serviceInfo.SID);
                }
            }
        }

        /// <summary>
        /// チャンネル一覧および予約一覧選択状態変更処理
        /// ヘッダを選択しようとした場合、選択状態を解除する。
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void serviceListView_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (e.IsSelected && string.IsNullOrEmpty(e.Item.Name))
            {
                e.Item.Selected = false;
            }
        }

        /// <summary>
        /// 予約一覧マウスクリック処理
        /// 右クリック時、対象の番組情報をコンテキストメニューに表示。
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void reserveListView_MouseClick(object sender, MouseEventArgs e)
        {
            var selectedCount = serviceListView.SelectedItems.Count;
            if (selectedCount == 0)
            {
                return;
            }

            // 右クリック
            // 対象の番組情報をコンテキストメニューで表示
            if (e.Button == MouseButtons.Right)
            {
                listContextMenuStrip.Items.Clear();

                try
                {
                    // クリックした箇所が自動選択されるので拾う
                    var selected = serviceListView.SelectedItems[0];

                    EpgEventInfo ev = allEventMap[selected.Name];

                    // Webリンク使用時のみ1行目にWebリンク用のボタンを表示
                    if (rockbarSetting.UseWebLink)
                    {
                        var item = listContextMenuStrip.Items.Add(">>");
                        item.Click += (s2, e2) => accessWebUrl(ev);

                        listContextMenuStrip.Items.Add(new ToolStripSeparator());
                    }

                    // 有効化・無効化を追加
                    var hasData = reserveMap.TryGetValue(selected.Name, out var reserveData);
                    if (hasData)
                    {
                        var item = listContextMenuStrip.Items.Add(reserveData.RecSetting.IsNoRec() ? "録画を有効にする" : "録画を無効にする");
                        item.Click += (s2, e2) => ToggleRecMode(reserveData);
                        listContextMenuStrip.Items.Add(new ToolStripSeparator());
                    }

                    var dateTime = $"{ev.start_time.ToString("yyyy/MM/dd(ddd) HH:mm")}-{ev.start_time.AddSeconds(ev.durationSec).ToString("HH:mm")}";
                    var dateItem = listContextMenuStrip.Items.Add(dateTime);
                    dateItem.Enabled = false;
                    listContextMenuStrip.Items.Add(new ToolStripSeparator());

                    // 予約情報をメニューに追加
                    if (hasData)
                    {
                        // 予約コメントを追加
                        if (! string.IsNullOrEmpty(reserveData.Comment))
                        {
                            var item = listContextMenuStrip.Items.Add(reserveData.Comment);
                            item.Enabled = false;
                            listContextMenuStrip.Items.Add(new ToolStripSeparator());
                        }
                    }

                    // 基本情報をメニューに追加
                    var shortStrs = RockbarUtility.BreakString(ev.ShortInfo?.text_char);

                    if (shortStrs != null)
                    {
                        foreach (string str in shortStrs)
                        {
                            var item = listContextMenuStrip.Items.Add(str);
                            item.Enabled = false;
                        }
                    }

                    listContextMenuStrip.Items.Add(new ToolStripSeparator());

                    // 拡張情報をメニューに追加
                    var longStrs = RockbarUtility.BreakString(ev.ExtInfo?.text_char);

                    if (longStrs != null)
                    {
                        foreach (string str in longStrs)
                        {
                            var item = listContextMenuStrip.Items.Add(str);
                            item.Enabled = false;
                        }
                    }
                }
                catch
                {
                    listContextMenuStrip.Items.Add("番組情報を取得できませんでした");
                }

                listContextMenuStrip.Show((Control)sender, new Point(e.X, e.Y));
            }
        }

        /// <summary>
        /// 予約一覧マウスアップ処理
        /// 中央クリック時、対象の予約情報の録画有効・無効を切り替える。
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void reserveListView_MouseUp(object sender, MouseEventArgs e)
        {
            // 中央クリック
            // 予約情報の録画有効・無効を切り替える
            if (e.Button == MouseButtons.Middle)
            {
                // クリックした箇所にある項目を取得する
                var selected = serviceListView.GetItemAt(e.Location.X, e.Location.Y);

                if (selected != null && reserveMap.TryGetValue(selected.Name, out var reserve))
                {
                    ToggleRecMode(reserve);
                }
            }
        }

        /// <summary>
        /// 予約一覧マウスダブルクリック処理
        /// Web LinkオプションがONの場合、カーソル箇所の番組のWeb番組情報にアクセスする。
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void reserveListView_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            // Web番組詳細使用時のみ
            if (! rockbarSetting.UseWebLink)
            {
                return;
            }

            var selectedCount = serviceListView.SelectedItems.Count;
            if (selectedCount == 0)
            {
                return;
            }

            // 左ダブルクリック
            // Web番組詳細にアクセスする
            if (e.Button == MouseButtons.Left)
            {
                try
                {
                    // クリックした箇所が自動選択されるので拾う
                    var selected = serviceListView.SelectedItems[0];

                    EpgEventInfo ev = allEventMap[selected.Name];
                    accessWebUrl(ev);
                }
                catch
                {
                    MessageBox.Show("番組情報を取得できませんでした。", "ブラウザ起動エラー");
                    return;
                }
            }
        }

        /// <summary>
        /// 録画済み情報マウスクリック処理
        /// 右クリック時、対象の番組情報をコンテキストメニューに表示。
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void recListView_MouseClick(object sender, MouseEventArgs e)
        {
            var selectedCount = serviceListView.SelectedItems.Count;
            if (selectedCount == 0)
            {
                return;
            }

            // 右クリック
            // 対象の番組情報をコンテキストメニューで表示
            if (e.Button == MouseButtons.Right)
            {
                listContextMenuStrip.Items.Clear();

                try
                {
                    // クリックした箇所が自動選択されるので拾う
                    var selected = serviceListView.SelectedItems[0];
                    uint recID;
                    if (! uint.TryParse(selected.Name, out recID))
                    {
                        MessageBox.Show("録画情報IDの取得に失敗しました。", "録画情報IDエラー");
                        return;
                    }

                    RecFileInfo recFile = recMap[recID];

                    // Webリンク使用時のみ1行目にWebリンク用のボタンを表示
                    if (rockbarSetting.UseWebLink)
                    {
                        var item = listContextMenuStrip.Items.Add(">>");
                        item.Click += (s2, e2) => accessWebUrl(recFile);

                        listContextMenuStrip.Items.Add(new ToolStripSeparator());
                    }

                    var detailTexts = createRecInfoDetailTexts(recFile);
                    if (detailTexts.Count > 0)
                    {
                        foreach (var texts in detailTexts.GetRange(0, detailTexts.Count - 1))
                        {
                            if (texts.Count == 0)
                            {
                                continue;
                            }
                            foreach (var text in texts)
                            {
                                var item = listContextMenuStrip.Items.Add(text);
                                item.Enabled = false;
                            }
                            listContextMenuStrip.Items.Add(new ToolStripSeparator());
                        }
                        var lastTexts = detailTexts[detailTexts.Count - 1];
                        if (lastTexts.Count > 0)
                        {
                            foreach (var text in lastTexts)
                            {
                                var item = listContextMenuStrip.Items.Add(text);
                                item.Enabled = false;
                            }
                        }
                    }
                }
                catch
                {
                    listContextMenuStrip.Items.Add("番組情報を取得できませんでした");
                }

                listContextMenuStrip.Show((Control)sender, new Point(e.X, e.Y));
            }
        }

        /// <summary>
        /// 録画済み情報マウスダブルクリック処理
        /// TVTest使用オプションがONの場合、カーソル箇所の番組を対象にTVTestを起動する。
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void recListView_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            // TVTest使用時のみ
            if (! rockbarSetting.UseDoubleClickTvtest)
            {
                return;
            }

            var selectedCount = serviceListView.SelectedItems.Count;
            if (selectedCount == 0)
            {
                return;
            }

            // 左ダブルクリック
            // TVTestを起動する
            if (e.Button == MouseButtons.Left)
            {
                // クリックした箇所が自動選択されるので拾う
                var selected = serviceListView.SelectedItems[0];

                uint recID;
                if (! uint.TryParse(selected.Name, out recID))
                {
                    MessageBox.Show("録画情報IDの取得に失敗しました。", "録画情報IDエラー");
                    return;
                }

                RecFileInfo recFile = recMap[recID];

                if (rockbarSetting.UseTcpIp && rockbarSetting.IpAddress.IndexOf("127.0.0.1") < 0)
                {
                    string networkPath = "";
                    ErrCode errCode = ctrlCmdUtil.SendGetRecFileNetworkPath(recFile.RecFilePath, ref networkPath);
                    if (errCode != ErrCode.CMD_SUCCESS || string.IsNullOrEmpty(networkPath))
                    {
                        MessageBox.Show("ネットワークパスの取得に失敗しました。EDCBの設定を見直してください。", "ネットワークパスエラー");
                        return;
                    }

                    // TvtPlay
                    startTvtPlay(networkPath);
                }
                else
                {
                    // TvtPlay
                    startTvtPlay(recFile.RecFilePath);
                }
            }
        }

        /// <summary>
        /// チューナー一覧マウスクリック時処理
        /// 右クリック時、直近30件の予約情報をコンテキストメニューに表示。
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void tunerListView_MouseClick(object sender, MouseEventArgs e)
        {
            // 右クリック
            // 今後の予約を30件までコンテキストメニューで表示
            //TODO DRY
            if (e.Button == MouseButtons.Right)
            {
                // クリックした箇所が自動選択されるので拾う
                var selected = tunerListView.SelectedItems[0];

                listContextMenuStrip.Items.Clear();

                TunerReserveInfo tunerReserveInfo = tunerReserveInfos.Find((TunerReserveInfo x) => x.tunerID.ToString() == selected.Name);

                // TunerReserveInfoには予約IDしか入っていないので、ReserveDataから予約情報を取り直す
                HashSet<uint> reserveIds = tunerReserveInfo.reserveList.ToHashSet();

                var reserves = reserveDatas.FindAll(x => reserveIds.Contains(x.ReserveID)).OrderBy(x => x.StartTime);

                int i = 0;

                foreach (var reserveData in reserves)
                {
                    string key = RockbarUtility.GetKey(reserveData.TransportStreamID, reserveData.ServiceID, reserveData.EventID);

                    EpgEventInfo ev = null;

                    if (allEventMap.ContainsKey(key))
                    {
                        ev = allEventMap[key];
                    }

                    listContextMenuStrip.Items.Add(createEventToolStripMenuItem(ev, reserveData, true));

                    i++;

                    if (i >= 30)
                    {
                        break;
                    }
                }

                listContextMenuStrip.Show((Control)sender, new Point(0, e.Y));
            }
        }

        /// <summary>
        /// タイマー処理
        /// 毎分0秒のEpgTimerSrv通信と、TVTestの起動・終了を行う
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void timer_Tick(object sender, EventArgs e)
        {
            DateTime timerTime = DateTime.Now;

            if (timerTime.Second == 0) {
                RefreshEvent(false, true);
            }
            
            if (! rockbarSetting.IsAutoOpenTvtest)
            {
                return;
            }

            // TVTest自動起動
            // 実実装としては毎秒チェックするのではなく、毎分(59-マージン)秒タイミングで次の1分間に始まる番組をオープンする
            if (timerTime.Second == (59 - rockbarSetting.AutoOpenMargin)) {
                // 0の場合はこの1分間なので59秒加算
                DateTime checkTime = timerTime.AddSeconds(59);

                // お気に入りサービスのキー(大した件数ではない想定なので毎秒計算し直しで良いものとする)
                HashSet<string> favoriteServiceKeys = favoriteServiceList.Select(x => RockbarUtility.GetKey(x.Tsid, x.Sid)).ToHashSet();

                foreach (var reserve in reserveDatas)
                {
                    if (reserve.StartTime.Date == checkTime.Date && reserve.StartTime.Hour == checkTime.Hour && reserve.StartTime.Minute == checkTime.Minute)
                    {
                        string key = RockbarUtility.GetKey(reserve.TransportStreamID, reserve.ServiceID);

                        // すでに自動起動中のTVTestとTSID・SIDが同一の場合、(TVTest側でチャンネルが変わっていない前提で)起動スキップする
                        if (tvtestProcesses.ContainsKey(key))
                        {
                            continue;
                        }

                        // お気に入りサービスオプションが設定されている場合、お気に入りサービスにキーが含まれていなかったらスキップ
                        if (rockbarSetting.IsAutoOpenTvtestFavoriteService && ! favoriteServiceKeys.Contains(key)) 
                        {
                            continue;
                        }

                        // 地上波
                        if (RockbarUtility.GetServiceType(reserve.OriginalNetworkID) == ServiceType.DTTV && rockbarSetting.IsAutoOpenTvtestDttv)
                        {
                            var p = startTvTest(true, reserve.TransportStreamID, reserve.ServiceID);
                            tvtestProcesses.Add(key, p);
                        }

                        // BS, CS
                        if (
                            RockbarUtility.GetServiceType(reserve.OriginalNetworkID) == ServiceType.BS && rockbarSetting.IsAutoOpenTvtestBs ||
                            RockbarUtility.GetServiceType(reserve.OriginalNetworkID) == ServiceType.CS && rockbarSetting.IsAutoOpenTvtestCs
                        )
                        {
                            var p = startTvTest(false, reserve.TransportStreamID, reserve.ServiceID);
                            tvtestProcesses.Add(key, p);
                        }
                    }
                }
            }

            // TVTest自動終了
            // 現時点のオプションにかかわらず、自身が開いたTVTestは予約終了時間でクローズ
            // 実実装としては毎秒チェックするのではなく、毎分マージン秒タイミングで現在放送してない番組をクローズ
            if (timerTime.Second == rockbarSetting.AutoCloseMargin)
            {
                // 閉じてるプロセスは取り除く
                var keys = tvtestProcesses.Keys.ToList();

                foreach (var key in keys)
                {
                    try
                    {
                        if (tvtestProcesses[key].HasExited)
                        {
                            tvtestProcesses.Remove(key);
                        }
                    }
                    catch
                    {
                        // 何らかの理由でプロセスにアクセスできない場合もキーを削除
                        tvtestProcesses.Remove(key);
                    }
                }

                // 把握してる生存プロセスの中で、録画放送に該当してるものがない場合はクローズ
                Dictionary<string, ReserveData> currentReserves = new Dictionary<string, ReserveData>();

                // 現在放送中番組を抽出
                foreach (var data in reserveDatas)
                {
                    // 途中処理があまりに遅いと、タイマー開始時とNowでズレが生じる可能性あり。問題がでたら検討
                    if (data.StartTime < DateTime.Now && data.StartTime.AddSeconds(data.DurationSecond) > DateTime.Now)
                    {
                        var key = RockbarUtility.GetKey(data.TransportStreamID, data.ServiceID);

                        // 予約方法次第で同一番組が二重に登録されているケースあり
                        if (! currentReserves.ContainsKey(key))
                        {
                            currentReserves.Add(key, data);
                        }
                    }
                }

                // 放送中番組でないTVTestを閉じる
                foreach (var p in tvtestProcesses)
                {
                    if (! currentReserves.ContainsKey(p.Key))
                    {
                        try
                        {
                            p.Value.CloseMainWindow();
                        }
                        catch
                        {
                            // 何らかの理由でプロセスがクローズできない場合
                            // 想定外ケースなのでここに落ちる場合は原因究明要
                        }
                    }
                }
            }
        }

        /// <summary>
        /// チャンネルタブ切り替え処理
        /// EpgTimerSrv通信せずに、対象チャンネルの表示切替を行う
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void serviceTabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshEvent(true, false);
        }

        /// <summary>
        /// ✕ボタン押下処理
        /// タスクトレイ格納オプションON時、タスクトレイに格納。そうでない場合、フォームを閉じる
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void closeButton_Click(object sender, EventArgs e)
        {
            if (rockbarSetting.StoreTaskTrayByClosing)
            {
                // 他のオプションにかかわらず、最小化(もどき)をする場合はタスクトレイにアイコンを表示する
                notifyIcon.Visible = true;
                this.Visible = false;
            }
            else
            {
                this.Close();
            }
        }

        /// <summary>
        /// メインフォームマウスボタン押下処理
        /// ドラッグ時の位置取得。
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void MainForm_MouseDown(object sender, MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Left) == MouseButtons.Left)
            {
                //位置を記憶する
                mousePoint = new Point(e.X, e.Y);
            }

            // フォームと同様にドラッグしたいラベルはEnabled = falseにしておく
        }

        /// <summary>
        /// メインフォームマウス移動処理
        /// ドラッグ時フォーム移動処理。
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void MainForm_MouseMove(object sender, MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Left) == MouseButtons.Left)
            {
                this.Left += e.X - mousePoint.X;
                this.Top += e.Y - mousePoint.Y;
            }
        }


        /// <summary>
        /// フィルタ入力欄キー押下処理
        /// Enter入力時にフィルタを行う。
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void filterTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Filter();
                e.SuppressKeyPress = true;
            }
        }

        /// <summary>
        /// リセットボタン押下処理
        /// フィルタ文字列をクリアしフィルタ結果をリセットする。
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void resetButton_Click(object sender, EventArgs e)
        {
            ResetFilter();
        }

        /// <summary>
        /// フィルタボタン押下処理
        /// フィルタを行う。
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void filterButton_Click(object sender, EventArgs e)
        {
            Filter();
        }

        /// <summary>
        /// 設定ボタン押下処理
        /// 設定フォームを開き、設定変更があった場合は設定を再読込して画面をリフレッシュする。
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void settingButton_Click(object sender, EventArgs e)
        {
            SettingForm settingForm = new SettingForm(ctrlCmdUtil, canConnect);
            DialogResult result = settingForm.ShowDialog();
            settingForm.Dispose();

            if (result == DialogResult.OK)
            {
                ReloadSetting();
                applySetting();
                RefreshEvent(true, true);
            }
        }

        /// <summary>
        /// フォームクローズ中処理
        /// フォームクローズ時に、位置・サイズ情報を設定ファイルに出力する。
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            rockbarSetting.X = this.Location.X;
            rockbarSetting.Y = this.Location.Y;
            rockbarSetting.Width = this.Size.Width;
            rockbarSetting.Height = this.Size.Height;
            rockbarSetting.SplitterDistance = splitContainer.SplitterDistance;

            Toml.WriteFile(rockbarSetting, RockbarUtility.GetTomlSettingFilePath());
        }

        /// <summary>
        /// タスクトレイ終了コンテキストメニュークリック処理
        /// フォームを閉じる。
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// タスクトレイアイコンマウスボタン押下処理
        /// トグルオプションONの場合、表示・非表示切り替え＋アクティブ化。それ以外の場合アクテイブ化のみ
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void notifyIcon_MouseDown(object sender, MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Left) == MouseButtons.Left)
            {
                if (this.Visible)
                {
                    // フォーム表示時に左クリックした場合はオプションにより挙動切り替え
                    if (rockbarSetting.ToggleVisibleTaskTrayIconClick)
                    {
                        this.Visible = false;
                    }
                    else
                    {
                        this.Activate();
                    }
                }
                else
                {
                    // フォーム非表示時に左クリックした場合は必ず表示
                    this.Visible = true;
                    this.Activate();

                    // オプションによりタスクトレイアイコン表示を切り替え
                    if (!rockbarSetting.ShowTaskTrayIcon)
                    {
                        notifyIcon.Visible = false;
                    }
                }
            }
        }

        /// <summary>
        /// フォームサイズ変更処理
        /// 左右のリストビューのカラム幅を調整する。
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void MainForm_SizeChanged(object sender, EventArgs e)
        {
            adjustListViewColumns(serviceListView);
            adjustListViewColumns(tunerListView);
        }

        /// <summary>
        /// スプリッタ位置調整処理
        /// 左右のリストビューのカラム幅を調整する。
        /// </summary>
        /// <param name="sender">イベントソース</param>
        /// <param name="e">イベントパラメータ</param>
        private void splitContainer_SplitterMoved(object sender, SplitterEventArgs e)
        {
            adjustListViewColumns(serviceListView);
            adjustListViewColumns(tunerListView);
        }
    }
}
