using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;

namespace AutoCAD_Block_Change
{
    public partial class BlockChangeForm : Form
    {
        private CheckBox chkReplaceBlocks;
        private CheckBox chkRenameBlocks;
        private CheckBox chkExportBlocks;
        
        private DataGridView dgvReplaceBlocks;
        private DataGridView dgvRenameBlocks;
        
        private TextBox txtExcelFileName;
        private TextBox txtExcelPath;
        private Button btnBrowsePath;
        
        private Button btnExecute;
        private Button btnCancel;
        
        private GroupBox grpReplace;
        private GroupBox grpRename;
        private GroupBox grpExport;

        // 新增匯入按鈕
        private Button btnImportReplace;
        private Button btnImportRename;
        private Button btnExportTemplate;

        // 新增圖層選擇控制項
        private CheckedListBox clbReplaceLayersFilter;
        private CheckedListBox clbRenameLayersFilter;
        private CheckBox chkFilterReplaceByLayer;
        private CheckBox chkFilterRenameByLayer;
        private Button btnRefreshReplaceLayer;
        private Button btnRefreshRenameLayer;

        // 修改：新增旋轉和聚合線相關控制項
        private CheckBox chkPreserveRotation;
        private CheckBox chkHandlePolylineConnection; // 修正：新增此控制項宣告
        
        // 新增：旋轉輸入控制項
        private CheckBox chkApplyAdditionalRotation;
        private RadioButton rbRotateLeft;
        private RadioButton rbRotateRight;
        private NumericUpDown nudRotationDegrees;
        private DataGridViewTextBoxColumn oldBlockCol;
        private DataGridViewTextBoxColumn newBlockCol;
        private DataGridViewTextBoxColumn oldNameCol;
        private DataGridViewTextBoxColumn newNameCol;
        private Label lblFileName;
        private Label lblPath;
        private Label lblRotationDegrees;

        // 定義現代化配色方案
        private static readonly Color PrimaryColor = Color.FromArgb(41, 128, 185);      // 主要藍色
        private static readonly Color SecondaryColor = Color.FromArgb(52, 152, 219);    // 次要藍色  
        private static readonly Color AccentColor = Color.FromArgb(46, 204, 113);       // 重點綠色
        private static readonly Color DangerColor = Color.FromArgb(231, 76, 60);        // 危險紅色
        private static readonly Color WarningColor = Color.FromArgb(243, 156, 18);      // 警告橙色
        private static readonly Color LightGray = Color.FromArgb(236, 240, 241);        // 淺灰色
        private static readonly Color DarkGray = Color.FromArgb(52, 73, 94);            // 深灰色
        private static readonly Color TextColor = Color.FromArgb(44, 62, 80);           // 文字顏色
        private static readonly Color CardBackground = Color.FromArgb(255, 255, 255);   // 卡片背景
        private static readonly Color FormBackground = Color.FromArgb(248, 249, 250);   // 表單背景

        public BlockChangeForm()
        {
            // 修正：初始化字符編碼支援
            InitializeEncodingSupport();
            InitializeComponent();
            ApplyModernStyling();
        }

        // 新增：初始化字符編碼支援
        private void InitializeEncodingSupport()
        {
            try
            {
                // 註冊編碼提供者以支援中文字符
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                
                // 設定控制台編碼為UTF-8
                Console.OutputEncoding = Encoding.UTF8;
                Console.InputEncoding = Encoding.UTF8;
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"編碼初始化失敗: {ex.Message}");
            }
        }

        private void ApplyModernStyling()
        {
            // 設定表單背景和字體
            this.BackColor = FormBackground;
            
            // 修正：改善中文字體支援
            try
            {
                // 優先使用 Microsoft YaHei，如果不存在則使用微軟正黑體或系統預設字體
                System.Drawing.Font baseFont = GetChineseFriendlyFont();
                this.Font = baseFont;
            }
            catch
            {
                // 如果字體設定失敗，使用系統預設字體
                this.Font = SystemFonts.DefaultFont;
            }
            
            // 設定標題文字樣式
            foreach (CheckBox chk in new[] { chkReplaceBlocks, chkRenameBlocks, chkExportBlocks })
            {
                StyleHeaderCheckBox(chk);
            }

            // 設定GroupBox樣式
            foreach (GroupBox grp in new[] { grpReplace, grpRename, grpExport })
            {
                StyleGroupBox(grp);
            }

            // 設定DataGridView樣式
            foreach (DataGridView dgv in new[] { dgvReplaceBlocks, dgvRenameBlocks })
            {
                StyleDataGridView(dgv);
            }

            // 設定CheckedListBox樣式
            foreach (CheckedListBox clb in new[] { clbReplaceLayersFilter, clbRenameLayersFilter })
            {
                StyleCheckedListBox(clb);
            }

            // 設定按鈕樣式
            StylePrimaryButton(btnExecute);
            StyleSecondaryButton(btnCancel);
            StyleAccentButton(btnImportReplace);
            StyleAccentButton(btnImportRename);
            StyleAccentButton(btnExportTemplate);
            StyleWarningButton(btnBrowsePath);
            StyleInfoButton(btnRefreshReplaceLayer);
            StyleInfoButton(btnRefreshRenameLayer);

            // 設定文字框樣式
            foreach (TextBox txt in new[] { txtExcelFileName, txtExcelPath })
            {
                StyleTextBox(txt);
            }

            // 設定標籤樣式
            foreach (Label lbl in new[] { lblFileName, lblPath, lblRotationDegrees })
            {
                StyleLabel(lbl);
            }

            // 設定CheckBox樣式
            foreach (CheckBox chk in new[] { chkFilterReplaceByLayer, chkFilterRenameByLayer, 
                chkPreserveRotation, chkHandlePolylineConnection, chkApplyAdditionalRotation })
            {
                StyleCheckBox(chk);
            }

            // 設定RadioButton樣式
            foreach (RadioButton rb in new[] { rbRotateLeft, rbRotateRight })
            {
                StyleRadioButton(rb);
            }

            // 設定NumericUpDown樣式
            StyleNumericUpDown(nudRotationDegrees);
        }

        // 新增：獲取中文友好字體
        private System.Drawing.Font GetChineseFriendlyFont()
        {
            // 嘗試不同的中文字體，按優先順序
            string[] fontNames = { 
                "Microsoft YaHei UI",  // 微軟雅黑 UI
                "Microsoft YaHei",     // 微軟雅黑
                "微軟正黑體",            // Microsoft JhengHei
                "Microsoft JhengHei UI",
                "Microsoft JhengHei",
                "SimSun",              // 宋體
                "SimHei",              // 黑體
                "NSimSun"              // 新宋體
            };

            foreach (string fontName in fontNames)
            {
                try
                {
                    using (var testFont = new System.Drawing.Font(fontName, 9F))
                    {
                        // 如果字體可以創建，返回該字體
                        return new System.Drawing.Font(fontName, 9F, FontStyle.Regular);
                    }
                }
                catch
                {
                    continue; // 嘗試下一個字體
                }
            }

            // 如果所有中文字體都失敗，返回系統預設字體
            return SystemFonts.DefaultFont;
        }

        private void StyleHeaderCheckBox(CheckBox chk)
        {
            chk.ForeColor = TextColor;
            chk.Font = new System.Drawing.Font(chk.Font.FontFamily, 12F, FontStyle.Bold);
            chk.BackColor = Color.Transparent;
        }

        private void StyleGroupBox(GroupBox grp)
        {
            grp.ForeColor = TextColor;
            grp.Font = new System.Drawing.Font(grp.Font.FontFamily, 10F, FontStyle.Bold);
            grp.BackColor = CardBackground;
            
            // 創建圓角邊框效果
            grp.Paint += (sender, e) =>
            {
                GroupBox gb = sender as GroupBox;
                Graphics g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                
                // 繪製圓角矩形背景
                using (GraphicsPath path = CreateRoundedRectangle(0, 0, gb.Width - 1, gb.Height - 1, 8))
                {
                    using (Brush brush = new SolidBrush(CardBackground))
                    {
                        g.FillPath(brush, path);
                    }
                    using (Pen pen = new Pen(LightGray, 2))
                    {
                        g.DrawPath(pen, path);
                    }
                }
            };
        }

        private void StyleDataGridView(DataGridView dgv)
        {
            dgv.BackgroundColor = CardBackground;
            dgv.GridColor = LightGray;
            dgv.BorderStyle = BorderStyle.None;
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv.MultiSelect = false;
            dgv.RowHeadersVisible = false;
            dgv.AllowUserToResizeRows = false;
            dgv.Font = new System.Drawing.Font(dgv.Font.FontFamily, 9F);
            
            // 標題樣式
            dgv.ColumnHeadersDefaultCellStyle.BackColor = PrimaryColor;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgv.Font.FontFamily, 9F, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Padding = new Padding(5);
            dgv.ColumnHeadersHeight = 40;
            
            // 資料列樣式
            dgv.DefaultCellStyle.BackColor = Color.White;
            dgv.DefaultCellStyle.ForeColor = TextColor;
            dgv.DefaultCellStyle.SelectionBackColor = SecondaryColor;
            dgv.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv.DefaultCellStyle.Padding = new Padding(5);
            dgv.RowTemplate.Height = 35;
            
            // 交替行顏色
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(249, 250, 251);
        }

        private void StyleCheckedListBox(CheckedListBox clb)
        {
            clb.BackColor = CardBackground;
            clb.ForeColor = TextColor;
            clb.BorderStyle = BorderStyle.None;
            clb.Font = new System.Drawing.Font(clb.Font.FontFamily, 9F);
            clb.ItemHeight = 24;
        }

        private void StylePrimaryButton(Button btn)
        {
            btn.BackColor = PrimaryColor;
            btn.ForeColor = Color.White;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.Font = new System.Drawing.Font(btn.Font.FontFamily, 10F, FontStyle.Bold);
            btn.Cursor = Cursors.Hand;
            
            AddButtonHoverEffect(btn, PrimaryColor, SecondaryColor);
            AddButtonRoundedCorners(btn);
        }

        private void StyleSecondaryButton(Button btn)
        {
            btn.BackColor = DarkGray;
            btn.ForeColor = Color.White;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.Font = new System.Drawing.Font(btn.Font.FontFamily, 10F, FontStyle.Bold);
            btn.Cursor = Cursors.Hand;
            
            AddButtonHoverEffect(btn, DarkGray, Color.FromArgb(70, 80, 90));
            AddButtonRoundedCorners(btn);
        }

        private void StyleAccentButton(Button btn)
        {
            btn.BackColor = AccentColor;
            btn.ForeColor = Color.White;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.Font = new System.Drawing.Font(btn.Font.FontFamily, 9F, FontStyle.Bold);
            btn.Cursor = Cursors.Hand;
            
            AddButtonHoverEffect(btn, AccentColor, Color.FromArgb(39, 174, 96));
            AddButtonRoundedCorners(btn);
        }

        private void StyleWarningButton(Button btn)
        {
            btn.BackColor = WarningColor;
            btn.ForeColor = Color.White;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.Font = new System.Drawing.Font(btn.Font.FontFamily, 9F, FontStyle.Bold);
            btn.Cursor = Cursors.Hand;
            
            AddButtonHoverEffect(btn, WarningColor, Color.FromArgb(211, 136, 15));
            AddButtonRoundedCorners(btn);
        }

        private void StyleInfoButton(Button btn)
        {
            btn.BackColor = SecondaryColor;
            btn.ForeColor = Color.White;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.Font = new System.Drawing.Font(btn.Font.FontFamily, 9F, FontStyle.Bold);
            btn.Cursor = Cursors.Hand;
            
            AddButtonHoverEffect(btn, SecondaryColor, PrimaryColor);
            AddButtonRoundedCorners(btn);
        }

        private void StyleTextBox(TextBox txt)
        {
            txt.BackColor = Color.White;
            txt.ForeColor = TextColor;
            txt.BorderStyle = BorderStyle.None;
            txt.Font = new System.Drawing.Font(txt.Font.FontFamily, 9F);
            
            // 創建圓角邊框效果
            Panel panel = new Panel();
            panel.BackColor = LightGray;
            panel.Size = new Size(txt.Width + 4, txt.Height + 4);
            panel.Location = new Point(txt.Location.X - 2, txt.Location.Y - 2);
            
            txt.Parent.Controls.Add(panel);
            panel.BringToFront();
            txt.Parent = panel;
            txt.Location = new Point(2, 2);
            
            panel.Paint += (sender, e) =>
            {
                Graphics g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (GraphicsPath path = CreateRoundedRectangle(0, 0, panel.Width - 1, panel.Height - 1, 4))
                {
                    using (Brush brush = new SolidBrush(Color.White))
                    {
                        g.FillPath(brush, path);
                    }
                    using (Pen pen = new Pen(LightGray, 1))
                    {
                        g.DrawPath(pen, path);
                    }
                }
            };
        }

        private void StyleLabel(Label lbl)
        {
            lbl.ForeColor = TextColor;
            lbl.Font = new System.Drawing.Font(lbl.Font.FontFamily, 9F, FontStyle.Regular);
            lbl.BackColor = Color.Transparent;
        }

        private void StyleCheckBox(CheckBox chk)
        {
            chk.ForeColor = TextColor;
            chk.Font = new System.Drawing.Font(chk.Font.FontFamily, 9F);
            chk.BackColor = Color.Transparent;
        }

        private void StyleRadioButton(RadioButton rb)
        {
            rb.ForeColor = TextColor;
            rb.Font = new System.Drawing.Font(rb.Font.FontFamily, 9F);
            rb.BackColor = Color.Transparent;
        }

        private void StyleNumericUpDown(NumericUpDown nud)
        {
            nud.BackColor = Color.White;
            nud.ForeColor = TextColor;
            nud.BorderStyle = BorderStyle.FixedSingle;
            nud.Font = new System.Drawing.Font(nud.Font.FontFamily, 9F);
        }

        private void AddButtonHoverEffect(Button btn, Color normalColor, Color hoverColor)
        {
            btn.MouseEnter += (sender, e) =>
            {
                btn.BackColor = hoverColor;
            };
            
            btn.MouseLeave += (sender, e) =>
            {
                btn.BackColor = normalColor;
            };
        }

        private void AddButtonRoundedCorners(Button btn)
        {
            btn.Paint += (sender, e) =>
            {
                Graphics g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                
                using (GraphicsPath path = CreateRoundedRectangle(0, 0, btn.Width - 1, btn.Height - 1, 6))
                {
                    using (Brush brush = new SolidBrush(btn.BackColor))
                    {
                        g.FillPath(brush, path);
                    }
                }
                
                // 繪製按鈕文字
                TextRenderer.DrawText(g, btn.Text, btn.Font, btn.ClientRectangle, btn.ForeColor,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            };
        }

        private GraphicsPath CreateRoundedRectangle(int x, int y, int width, int height, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            path.AddArc(x, y, radius * 2, radius * 2, 180, 90);
            path.AddArc(x + width - radius * 2, y, radius * 2, radius * 2, 270, 90);
            path.AddArc(x + width - radius * 2, y + height - radius * 2, radius * 2, radius * 2, 0, 90);
            path.AddArc(x, y + height - radius * 2, radius * 2, radius * 2, 90, 90);
            path.CloseAllFigures();
            return path;
        }

        private void InitializeComponent()
        {
            chkReplaceBlocks = new CheckBox();
            chkRenameBlocks = new CheckBox();
            chkExportBlocks = new CheckBox();
            grpReplace = new GroupBox();
            dgvReplaceBlocks = new DataGridView();
            oldBlockCol = new DataGridViewTextBoxColumn();
            newBlockCol = new DataGridViewTextBoxColumn();
            btnImportReplace = new Button();
            chkFilterReplaceByLayer = new CheckBox();
            btnRefreshReplaceLayer = new Button();
            clbReplaceLayersFilter = new CheckedListBox();
            chkPreserveRotation = new CheckBox();
            chkHandlePolylineConnection = new CheckBox(); // 修正：確保此控制項正確初始化
            chkApplyAdditionalRotation = new CheckBox();
            rbRotateLeft = new RadioButton();
            rbRotateRight = new RadioButton();
            lblRotationDegrees = new Label();
            nudRotationDegrees = new NumericUpDown();
            grpRename = new GroupBox();
            dgvRenameBlocks = new DataGridView();
            oldNameCol = new DataGridViewTextBoxColumn();
            newNameCol = new DataGridViewTextBoxColumn();
            btnImportRename = new Button();
            chkFilterRenameByLayer = new CheckBox();
            btnRefreshRenameLayer = new Button();
            clbRenameLayersFilter = new CheckedListBox();
            grpExport = new GroupBox();
            lblFileName = new Label();
            txtExcelFileName = new TextBox();
            lblPath = new Label();
            txtExcelPath = new TextBox();
            btnBrowsePath = new Button();
            btnExportTemplate = new Button();
            btnExecute = new Button();
            btnCancel = new Button();
            grpReplace.SuspendLayout();
            ((ISupportInitialize)dgvReplaceBlocks).BeginInit();
            ((ISupportInitialize)nudRotationDegrees).BeginInit();
            grpRename.SuspendLayout();
            ((ISupportInitialize)dgvRenameBlocks).BeginInit();
            grpExport.SuspendLayout();
            SuspendLayout();
            
            // Form 基本設定
            this.ClientSize = new Size(2200, 1768);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "BlockChangeForm";
            this.Padding = new Padding(20);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "圖塊管理工具";
            this.BackColor = FormBackground;
            
            // 主要功能選擇區域 - 頂部
            chkReplaceBlocks.Location = new Point(30, 30);
            chkReplaceBlocks.Name = "chkReplaceBlocks";
            chkReplaceBlocks.Size = new Size(240, 44);
            chkReplaceBlocks.TabIndex = 0;
            chkReplaceBlocks.Text = "批量置換圖塊";
            chkReplaceBlocks.UseVisualStyleBackColor = true;
            chkReplaceBlocks.CheckedChanged += ChkReplaceBlocks_CheckedChanged;

            chkRenameBlocks.Location = new Point(287, 30);
            chkRenameBlocks.Name = "chkRenameBlocks";
            chkRenameBlocks.Size = new Size(275, 44);
            chkRenameBlocks.TabIndex = 1;
            chkRenameBlocks.Text = "批量重命名圖塊";
            chkRenameBlocks.UseVisualStyleBackColor = true;
            chkRenameBlocks.CheckedChanged += ChkRenameBlocks_CheckedChanged;

            chkExportBlocks.Location = new Point(568, 28);
            chkExportBlocks.Name = "chkExportBlocks";
            chkExportBlocks.Size = new Size(238, 49);
            chkExportBlocks.TabIndex = 2;
            chkExportBlocks.Text = "輸出圖塊清單";
            chkExportBlocks.UseVisualStyleBackColor = true;
            chkExportBlocks.CheckedChanged += ChkExportBlocks_CheckedChanged;

            // 置換圖塊區域
            grpReplace.Controls.Add(dgvReplaceBlocks);
            grpReplace.Controls.Add(btnImportReplace);
            grpReplace.Controls.Add(chkFilterReplaceByLayer);
            grpReplace.Controls.Add(btnRefreshReplaceLayer);
            grpReplace.Controls.Add(clbReplaceLayersFilter);
            grpReplace.Controls.Add(chkPreserveRotation);
            grpReplace.Controls.Add(chkHandlePolylineConnection); // 修正：確保加入控制項
            grpReplace.Controls.Add(chkApplyAdditionalRotation);
            grpReplace.Controls.Add(rbRotateLeft);
            grpReplace.Controls.Add(rbRotateRight);
            grpReplace.Controls.Add(lblRotationDegrees);
            grpReplace.Controls.Add(nudRotationDegrees);
            grpReplace.Enabled = false;
            grpReplace.Location = new Point(30, 83);
            grpReplace.Name = "grpReplace";
            grpReplace.Size = new Size(2140, 622);
            grpReplace.TabIndex = 3;
            grpReplace.TabStop = false;
            grpReplace.Text = "置換圖塊設定";

            // DataGridView - 置換清單
            dgvReplaceBlocks.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvReplaceBlocks.Columns.AddRange(new DataGridViewColumn[] { oldBlockCol, newBlockCol });
            dgvReplaceBlocks.Location = new Point(20, 40);
            dgvReplaceBlocks.Name = "dgvReplaceBlocks";
            dgvReplaceBlocks.RowHeadersWidth = 102;
            dgvReplaceBlocks.Size = new Size(823, 565);
            dgvReplaceBlocks.TabIndex = 0;
            dgvReplaceBlocks.AllowUserToAddRows = true;
            dgvReplaceBlocks.AllowUserToDeleteRows = true;

            oldBlockCol.HeaderText = "舊圖塊名稱";
            oldBlockCol.Name = "oldBlockCol";
            oldBlockCol.Width = 340;

            newBlockCol.HeaderText = "新圖塊名稱";
            newBlockCol.Name = "newBlockCol";
            newBlockCol.Width = 340;

            // 按鈕區域 - 置換
            btnImportReplace.Location = new Point(1048, 491);
            btnImportReplace.Name = "btnImportReplace";
            btnImportReplace.Size = new Size(220, 125);
            btnImportReplace.TabIndex = 1;
            btnImportReplace.Text = "匯入CSV";
            btnImportReplace.UseVisualStyleBackColor = true;
            btnImportReplace.Click += BtnImportReplace_Click;

            // 圖層篩選區域 - 置換
            chkFilterReplaceByLayer.Location = new Point(1274, 30);
            chkFilterReplaceByLayer.Name = "chkFilterReplaceByLayer";
            chkFilterReplaceByLayer.Size = new Size(214, 50);
            chkFilterReplaceByLayer.TabIndex = 2;
            chkFilterReplaceByLayer.Text = "按圖層篩選";
            chkFilterReplaceByLayer.UseVisualStyleBackColor = true;
            chkFilterReplaceByLayer.CheckedChanged += ChkFilterReplaceByLayer_CheckedChanged;

            btnRefreshReplaceLayer.Location = new Point(1490, 30);
            btnRefreshReplaceLayer.Name = "btnRefreshReplaceLayer";
            btnRefreshReplaceLayer.Size = new Size(212, 60);
            btnRefreshReplaceLayer.TabIndex = 3;
            btnRefreshReplaceLayer.Text = "刷新";
            btnRefreshReplaceLayer.UseVisualStyleBackColor = true;
            btnRefreshReplaceLayer.Click += BtnRefreshReplaceLayer_Click;

            clbReplaceLayersFilter.CheckOnClick = true;
            clbReplaceLayersFilter.Enabled = false;
            clbReplaceLayersFilter.Location = new Point(1274, 96);
            clbReplaceLayersFilter.Name = "clbReplaceLayersFilter";
            clbReplaceLayersFilter.Size = new Size(846, 520);
            clbReplaceLayersFilter.TabIndex = 4;

            // 旋轉控制區域
            chkPreserveRotation.Checked = true;
            chkPreserveRotation.CheckState = CheckState.Checked;
            chkPreserveRotation.Location = new Point(855, 45);
            chkPreserveRotation.Name = "chkPreserveRotation";
            chkPreserveRotation.Size = new Size(302, 50);
            chkPreserveRotation.TabIndex = 5;
            chkPreserveRotation.Text = "保持圖塊旋轉角度";
            chkPreserveRotation.UseVisualStyleBackColor = true;

            // 修正：加入聚合線處理選項
            chkHandlePolylineConnection.Location = new Point(855, 101);
            chkHandlePolylineConnection.Name = "chkHandlePolylineConnection";
            chkHandlePolylineConnection.Size = new Size(200, 30);
            chkHandlePolylineConnection.TabIndex = 6;
            chkHandlePolylineConnection.Text = "智能聚合線處理";
            chkHandlePolylineConnection.UseVisualStyleBackColor = true;

            chkApplyAdditionalRotation.Location = new Point(855, 137);
            chkApplyAdditionalRotation.Name = "chkApplyAdditionalRotation";
            chkApplyAdditionalRotation.Size = new Size(190, 42);
            chkApplyAdditionalRotation.TabIndex = 7;
            chkApplyAdditionalRotation.Text = "額外旋轉";
            chkApplyAdditionalRotation.UseVisualStyleBackColor = true;
            chkApplyAdditionalRotation.CheckedChanged += ChkApplyAdditionalRotation_CheckedChanged;

            rbRotateLeft.Checked = true;
            rbRotateLeft.Enabled = false;
            rbRotateLeft.Location = new Point(895, 185);
            rbRotateLeft.Name = "rbRotateLeft";
            rbRotateLeft.Size = new Size(199, 48);
            rbRotateLeft.TabIndex = 8;
            rbRotateLeft.TabStop = true;
            rbRotateLeft.Text = "向左旋轉";
            rbRotateLeft.UseVisualStyleBackColor = true;

            rbRotateRight.Enabled = false;
            rbRotateRight.Location = new Point(1100, 181);
            rbRotateRight.Name = "rbRotateRight";
            rbRotateRight.Size = new Size(190, 52);
            rbRotateRight.TabIndex = 9;
            rbRotateRight.Text = "向右旋轉";
            rbRotateRight.UseVisualStyleBackColor = true;

            lblRotationDegrees.Location = new Point(941, 247);
            lblRotationDegrees.Name = "lblRotationDegrees";
            lblRotationDegrees.Size = new Size(145, 43);
            lblRotationDegrees.TabIndex = 10;
            lblRotationDegrees.Text = "旋轉度數:";

            nudRotationDegrees.DecimalPlaces = 1;
            nudRotationDegrees.Enabled = false;
            nudRotationDegrees.Location = new Point(1092, 247);
            nudRotationDegrees.Maximum = new decimal(new int[] { 360, 0, 0, 0 });
            nudRotationDegrees.Name = "nudRotationDegrees";
            nudRotationDegrees.Size = new Size(172, 46);
            nudRotationDegrees.TabIndex = 11;
            nudRotationDegrees.Value = new decimal(new int[] { 90, 0, 0, 0 });

            // 重命名圖塊區域
            grpRename.Controls.Add(dgvRenameBlocks);
            grpRename.Controls.Add(btnImportRename);
            grpRename.Controls.Add(chkFilterRenameByLayer);
            grpRename.Controls.Add(btnRefreshRenameLayer);
            grpRename.Controls.Add(clbRenameLayersFilter);
            grpRename.Enabled = false;
            grpRename.Location = new Point(30, 711);
            grpRename.Name = "grpRename";
            grpRename.Size = new Size(2140, 652);
            grpRename.TabIndex = 4;
            grpRename.TabStop = false;
            grpRename.Text = "重命名圖塊設定";

            // DataGridView - 重命名清單
            dgvRenameBlocks.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvRenameBlocks.Columns.AddRange(new DataGridViewColumn[] { oldNameCol, newNameCol });
            dgvRenameBlocks.Location = new Point(20, 40);
            dgvRenameBlocks.Name = "dgvRenameBlocks";
            dgvRenameBlocks.RowHeadersWidth = 102;
            dgvRenameBlocks.Size = new Size(1022, 596);
            dgvRenameBlocks.TabIndex = 0;
            dgvRenameBlocks.AllowUserToAddRows = true;
            dgvRenameBlocks.AllowUserToDeleteRows = true;

            oldNameCol.HeaderText = "舊圖塊名稱";
            oldNameCol.Name = "oldNameCol";
            oldNameCol.Width = 340;

            newNameCol.HeaderText = "新圖塊名稱";
            newNameCol.Name = "newNameCol";
            newNameCol.Width = 340;

            btnImportRename.Location = new Point(1048, 474);
            btnImportRename.Name = "btnImportRename";
            btnImportRename.Size = new Size(220, 162);
            btnImportRename.TabIndex = 1;
            btnImportRename.Text = "匯入CSV";
            btnImportRename.UseVisualStyleBackColor = true;
            btnImportRename.Click += BtnImportRename_Click;

            // 圖層篩選區域 - 重命名
            chkFilterRenameByLayer.Location = new Point(1274, 40);
            chkFilterRenameByLayer.Name = "chkFilterRenameByLayer";
            chkFilterRenameByLayer.Size = new Size(210, 58);
            chkFilterRenameByLayer.TabIndex = 2;
            chkFilterRenameByLayer.Text = "按圖層篩選";
            chkFilterRenameByLayer.UseVisualStyleBackColor = true;
            chkFilterRenameByLayer.CheckedChanged += ChkFilterRenameByLayer_CheckedChanged;

            btnRefreshRenameLayer.Location = new Point(1490, 40);
            btnRefreshRenameLayer.Name = "btnRefreshRenameLayer";
            btnRefreshRenameLayer.Size = new Size(212, 58);
            btnRefreshRenameLayer.TabIndex = 3;
            btnRefreshRenameLayer.Text = "刷新";
            btnRefreshRenameLayer.UseVisualStyleBackColor = true;
            btnRefreshRenameLayer.Click += BtnRefreshRenameLayer_Click;

            clbRenameLayersFilter.CheckOnClick = true;
            clbRenameLayersFilter.Enabled = false;
            clbRenameLayersFilter.Location = new Point(1274, 116);
            clbRenameLayersFilter.Name = "clbRenameLayersFilter";
            clbRenameLayersFilter.Size = new Size(846, 520);
            clbRenameLayersFilter.TabIndex = 4;

            // 輸出設定區域
            grpExport.Controls.Add(lblFileName);
            grpExport.Controls.Add(txtExcelFileName);
            grpExport.Controls.Add(lblPath);
            grpExport.Controls.Add(txtExcelPath);
            grpExport.Controls.Add(btnBrowsePath);
            grpExport.Controls.Add(btnExportTemplate);
            grpExport.Enabled = false;
            grpExport.Location = new Point(23, 1505);
            grpExport.Name = "grpExport";
            grpExport.Size = new Size(1297, 240);
            grpExport.TabIndex = 5;
            grpExport.TabStop = false;
            grpExport.Text = "輸出CSV設定";

            lblFileName.Location = new Point(16, 81);
            lblFileName.Name = "lblFileName";
            lblFileName.Size = new Size(146, 45);
            lblFileName.TabIndex = 0;
            lblFileName.Text = "檔案名稱:";

            txtExcelFileName.Location = new Point(168, 81);
            txtExcelFileName.Name = "txtExcelFileName";
            txtExcelFileName.Size = new Size(529, 46);
            txtExcelFileName.TabIndex = 1;
            txtExcelFileName.Text = "圖塊清單";

            lblPath.Location = new Point(16, 159);
            lblPath.Name = "lblPath";
            lblPath.Size = new Size(146, 48);
            lblPath.TabIndex = 2;
            lblPath.Text = "存放路徑:";

            txtExcelPath.Location = new Point(168, 156);
            txtExcelPath.Name = "txtExcelPath";
            txtExcelPath.Size = new Size(529, 46);
            txtExcelPath.TabIndex = 3;
            txtExcelPath.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            btnBrowsePath.Location = new Point(703, 154);
            btnBrowsePath.Name = "btnBrowsePath";
            btnBrowsePath.Size = new Size(147, 48);
            btnBrowsePath.TabIndex = 4;
            btnBrowsePath.Text = "瀏覽";
            btnBrowsePath.UseVisualStyleBackColor = true;
            btnBrowsePath.Click += BtnBrowsePath_Click;

            btnExportTemplate.Location = new Point(1055, 45);
            btnExportTemplate.Name = "btnExportTemplate";
            btnExportTemplate.Size = new Size(220, 156);
            btnExportTemplate.TabIndex = 5;
            btnExportTemplate.Text = "匯出範本";
            btnExportTemplate.UseVisualStyleBackColor = true;
            btnExportTemplate.Click += BtnExportTemplate_Click;

            // 主要操作按鈕
            btnExecute.Location = new Point(1850, 1470);
            btnExecute.Name = "btnExecute";
            btnExecute.Size = new Size(150, 50);
            btnExecute.TabIndex = 6;
            btnExecute.Text = "執行操作";
            btnExecute.UseVisualStyleBackColor = true;
            btnExecute.Click += BtnExecute_Click;

            btnCancel.DialogResult = DialogResult.Cancel;
            btnCancel.Location = new Point(2020, 1470);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new Size(120, 50);
            btnCancel.TabIndex = 7;
            btnCancel.Text = "取消";
            btnCancel.UseVisualStyleBackColor = true;

            // 加入所有控制項到表單
            this.Controls.AddRange(new Control[] {
                chkReplaceBlocks, chkRenameBlocks, chkExportBlocks,
                grpReplace, grpRename, grpExport,
                btnExecute, btnCancel
            });

            grpReplace.ResumeLayout(false);
            ((ISupportInitialize)dgvReplaceBlocks).EndInit();
            ((ISupportInitialize)nudRotationDegrees).EndInit();
            grpRename.ResumeLayout(false);
            ((ISupportInitialize)dgvRenameBlocks).EndInit();
            grpExport.ResumeLayout(false);
            grpExport.PerformLayout();
            ResumeLayout(false);
        }

        private void ChkApplyAdditionalRotation_CheckedChanged(object sender, EventArgs e)
        {
            bool enabled = chkApplyAdditionalRotation.Checked;
            rbRotateLeft.Enabled = enabled;
            rbRotateRight.Enabled = enabled;
            nudRotationDegrees.Enabled = enabled;
        }

        private void ChkReplaceBlocks_CheckedChanged(object sender, EventArgs e)
        {
            grpReplace.Enabled = chkReplaceBlocks.Checked;
            if (chkReplaceBlocks.Checked)
            {
                LoadLayers(clbReplaceLayersFilter);
            }
        }

        private void ChkRenameBlocks_CheckedChanged(object sender, EventArgs e)
        {
            grpRename.Enabled = chkRenameBlocks.Checked;
            if (chkRenameBlocks.Checked)
            {
                LoadLayers(clbRenameLayersFilter);
            }
        }

        private void ChkExportBlocks_CheckedChanged(object sender, EventArgs e)
        {
            grpExport.Enabled = chkExportBlocks.Checked;
        }

        private void ChkFilterReplaceByLayer_CheckedChanged(object sender, EventArgs e)
        {
            clbReplaceLayersFilter.Enabled = chkFilterReplaceByLayer.Checked;
            if (chkFilterReplaceByLayer.Checked && clbReplaceLayersFilter.Items.Count == 0)
            {
                LoadLayers(clbReplaceLayersFilter);
            }
        }

        private void ChkFilterRenameByLayer_CheckedChanged(object sender, EventArgs e)
        {
            clbRenameLayersFilter.Enabled = chkFilterRenameByLayer.Checked;
            if (chkFilterRenameByLayer.Checked && clbRenameLayersFilter.Items.Count == 0)
            {
                LoadLayers(clbRenameLayersFilter);
            }
        }

        private void BtnRefreshReplaceLayer_Click(object sender, EventArgs e)
        {
            LoadLayers(clbReplaceLayersFilter);
        }

        private void BtnRefreshRenameLayer_Click(object sender, EventArgs e)
        {
            LoadLayers(clbRenameLayersFilter);
        }

        private void LoadLayers(CheckedListBox layerListBox)
        {
            try
            {
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return;

                layerListBox.Items.Clear();

                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    LayerTable layerTable = tr.GetObject(doc.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                    
                    var layerNames = new List<string>();
                    foreach (ObjectId layerId in layerTable)
                    {
                        LayerTableRecord layer = tr.GetObject(layerId, OpenMode.ForRead) as LayerTableRecord;
                        layerNames.Add(layer.Name);
                    }

                    // 排序圖層名稱
                    layerNames.Sort();
                    
                    foreach (string layerName in layerNames)
                    {
                        layerListBox.Items.Add(layerName);
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"載入圖層失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private HashSet<string> GetSelectedLayers(CheckedListBox layerListBox, bool isFilterEnabled)
        {
            var selectedLayers = new HashSet<string>();
            
            if (isFilterEnabled)
            {
                foreach (string layerName in layerListBox.CheckedItems)
                {
                    selectedLayers.Add(layerName);
                }
            }
            
            return selectedLayers;
        }

        private void BtnBrowsePath_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.SelectedPath = txtExcelPath.Text;
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtExcelPath.Text = dialog.SelectedPath;
                }
            }
        }

        private void BtnImportReplace_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog dialog = new OpenFileDialog())
                {
                    dialog.Filter = "CSV 檔案|*.csv|所有檔案|*.*";
                    dialog.Title = "選擇置換圖塊CSV檔案";
                    
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        ImportCsvToGrid(dialog.FileName, dgvReplaceBlocks);
                        MessageBox.Show("置換圖塊數據匯入成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"匯入失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnImportRename_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog dialog = new OpenFileDialog())
                {
                    dialog.Filter = "CSV 檔案|*.csv|所有檔案|*.*";
                    dialog.Title = "選擇重命名圖塊CSV檔案";
                    
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        ImportCsvToGrid(dialog.FileName, dgvRenameBlocks);
                        MessageBox.Show("重命名圖塊數據匯入成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"匯入失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnExportTemplate_Click(object sender, EventArgs e)
        {
            try
            {
                using (SaveFileDialog dialog = new SaveFileDialog())
                {
                    dialog.Filter = "CSV 檔案|*.csv";
                    dialog.Title = "匯出CSV範本";
                    dialog.FileName = "圖塊操作範本.csv";
                    
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        // 創建範本 - 修正：使用UTF-8 BOM確保中文正確顯示
                        var templateContent = new StringBuilder();
                        templateContent.AppendLine("舊圖塊名稱,新圖塊名稱");
                        templateContent.AppendLine("DOOR_OLD,DOOR_NEW");
                        templateContent.AppendLine("WINDOW_V1,WINDOW_V2");
                        templateContent.AppendLine("BLOCK1,標準門");
                        
                        // 修正：使用UTF-8 BOM編碼寫入文件
                        var utf8WithBom = new UTF8Encoding(true);
                        File.WriteAllText(dialog.FileName, templateContent.ToString(), utf8WithBom);
                        MessageBox.Show("範本匯出成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"匯出範本失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ImportCsvToGrid(string filePath, DataGridView grid)
        {
            // 修正：使用UTF-8編碼讀取CSV文件
            var lines = File.ReadAllLines(filePath, Encoding.UTF8);
            
            // 清空現有數據
            grid.Rows.Clear();
            
            // 跳過標題行（如果存在）
            int startIndex = 0;
            if (lines.Length > 0 && (lines[0].Contains("舊") || lines[0].Contains("Old")))
            {
                startIndex = 1;
            }
            
            // 匯入數據
            for (int i = startIndex; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;
                
                var parts = lines[i].Split(',');
                if (parts.Length >= 2)
                {
                    int rowIndex = grid.Rows.Add();
                    grid.Rows[rowIndex].Cells[0].Value = parts[0].Trim();
                    grid.Rows[rowIndex].Cells[1].Value = parts[1].Trim();
                }
            }
        }

        private void BtnExecute_Click(object sender, EventArgs e)
        {
            try
            {
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("沒有活動的 AutoCAD 文檔。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (chkReplaceBlocks.Checked)
                {
                    ExecuteReplaceBlocks(doc);
                }

                if (chkRenameBlocks.Checked)
                {
                    ExecuteRenameBlocks(doc);
                }

                if (chkExportBlocks.Checked)
                {
                    ExecuteExportBlocks(doc);
                }

                MessageBox.Show("操作完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.DialogResult = DialogResult.OK;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"執行時發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExecuteReplaceBlocks(Document doc)
        {
            Database db = doc.Database;
            var selectedLayers = GetSelectedLayers(clbReplaceLayersFilter, chkFilterReplaceByLayer.Checked);
            
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                int totalReplacedCount = 0;
                int polylineAdjustedCount = 0;

                foreach (DataGridViewRow row in dgvReplaceBlocks.Rows)
                {
                    if (row.IsNewRow) continue;

                    string oldBlockName = row.Cells["oldBlockCol"].Value?.ToString();
                    string newBlockName = row.Cells["newBlockCol"].Value?.ToString();

                    if (string.IsNullOrEmpty(oldBlockName) || string.IsNullOrEmpty(newBlockName))
                        continue;

                    if (!bt.Has(oldBlockName) || !bt.Has(newBlockName))
                        continue;

                    ObjectId newBlockId = bt[newBlockName];
                    var blocksToReplace = new List<BlockReplaceInfo>();

                    foreach (ObjectId objId in modelSpace)
                    {
                        if (objId.ObjectClass == RXClass.GetClass(typeof(BlockReference)))
                        {
                            BlockReference blockRef = tr.GetObject(objId, OpenMode.ForRead) as BlockReference;
                            if (blockRef.Name.Equals(oldBlockName, StringComparison.OrdinalIgnoreCase))
                            {
                                // 檢查圖層篩選
                                if (chkFilterReplaceByLayer.Checked && selectedLayers.Count > 0)
                                {
                                    if (!selectedLayers.Contains(blockRef.Layer))
                                        continue;
                                }

                                var replaceInfo = new BlockReplaceInfo
                                {
                                    ObjectId = objId,
                                    Transform = blockRef.BlockTransform,
                                    Layer = blockRef.Layer,
                                    Position = blockRef.Position,
                                    Rotation = blockRef.Rotation,
                                    ScaleFactors = blockRef.ScaleFactors
                                };

                                // 如果啟用聚合線處理，尋找連接的聚合線
                                if (chkHandlePolylineConnection.Checked)
                                {
                                    replaceInfo.ConnectedPolylines = FindConnectedPolylines(tr, blockRef.Position);
                                }

                                blocksToReplace.Add(replaceInfo);
                            }
                        }
                    }

                    foreach (var replaceInfo in blocksToReplace)
                    {
                        // 刪除舊圖塊
                        Entity oldBlock = tr.GetObject(replaceInfo.ObjectId, OpenMode.ForWrite) as Entity;
                        oldBlock.Erase();

                        // 創建新圖塊參照
                        BlockReference newBlockRef = new BlockReference(Point3d.Origin, newBlockId);
                        
                        Point3d finalPosition = replaceInfo.Position;

                        // 計算最終旋轉角度
                        double finalRotation = 0;
                        if (chkPreserveRotation.Checked)
                        {
                            finalRotation = replaceInfo.Rotation;
                        }

                        // 新增：應用額外旋轉
                        if (chkApplyAdditionalRotation.Checked)
                        {
                            double additionalRotation = (double)nudRotationDegrees.Value * Math.PI / 180.0; // 轉換為弧度
                            
                            if (rbRotateLeft.Checked)
                            {
                                finalRotation += additionalRotation; // 向左旋轉（逆時針）
                            }
                            else if (rbRotateRight.Checked)
                            {
                                finalRotation -= additionalRotation; // 向右旋轉（順時針）
                            }
                        }

                        // 應用變換
                        newBlockRef.TransformBy(Matrix3d.Displacement(finalPosition - Point3d.Origin));
                        newBlockRef.Rotation = finalRotation;
                        newBlockRef.ScaleFactors = replaceInfo.ScaleFactors;
                        newBlockRef.Layer = replaceInfo.Layer;

                        modelSpace.AppendEntity(newBlockRef);
                        tr.AddNewlyCreatedDBObject(newBlockRef, true);

                        // 處理聚合線連接（智能延伸/裁剪考慮圖塊邊界）
                        if (chkHandlePolylineConnection.Checked && replaceInfo.ConnectedPolylines.Count > 0)
                        {
                            SmartExtendOrTrimPolylines(tr, replaceInfo.ConnectedPolylines, finalPosition, newBlockRef);
                            polylineAdjustedCount += replaceInfo.ConnectedPolylines.Count;
                        }

                        totalReplacedCount++;
                    }
                }

                tr.Commit();

                // 顯示執行結果
                var resultMessage = $"成功替換了 {totalReplacedCount} 個圖塊";
                if (chkHandlePolylineConnection.Checked && polylineAdjustedCount > 0)
                {
                    resultMessage += $"\n智能處理了 {polylineAdjustedCount} 個聚合線連接";
                }
                
                MessageBox.Show(resultMessage, "置換完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // 儲存圖塊置換信息的類別
        private class BlockReplaceInfo
        {
            public ObjectId ObjectId { get; set; }
            public Matrix3d Transform { get; set; }
            public string Layer { get; set; }
            public Point3d Position { get; set; }
            public double Rotation { get; set; }
            public Scale3d ScaleFactors { get; set; }
            public List<PolylineConnectionInfo> ConnectedPolylines { get; set; } = new List<PolylineConnectionInfo>();
        }

        // 更新：聚合線連接信息類別
        private class PolylineConnectionInfo
        {
            public ObjectId PolylineId { get; set; }
            public int VertexIndex { get; set; }
            public Point3d ConnectionPoint { get; set; }
            public double Distance { get; set; }
            public bool IsStartPoint { get; set; }
            public bool IsEndPoint { get; set; }
        }

        // 修改：尋找與特定圖塊直接連接的聚合線（改善容差和檢測邏輯）
        private List<PolylineConnectionInfo> FindConnectedPolylines(Transaction tr, Point3d blockPosition, double tolerance = 2.0)
        {
            var connections = new List<PolylineConnectionInfo>();
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            
            BlockTableRecord modelSpace = tr.GetObject(
                ((BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead))[BlockTableRecord.ModelSpace], 
                OpenMode.ForRead) as BlockTableRecord;

            int polylineCount = 0;
            int connectionCount = 0;

            foreach (ObjectId objId in modelSpace)
            {
                if (objId.ObjectClass == RXClass.GetClass(typeof(Polyline)))
                {
                    polylineCount++;
                    Polyline pline = tr.GetObject(objId, OpenMode.ForRead) as Polyline;
                    
                    // 檢查聚合線的端點是否在容差範圍內
                    for (int i = 0; i < pline.NumberOfVertices; i++)
                    {
                        Point3d vertex = pline.GetPoint3dAt(i);
                        double distance = vertex.DistanceTo(blockPosition);
                        
                        // 檢查是否為端點且在容差範圍內
                        if ((i == 0 || i == pline.NumberOfVertices - 1) && distance <= tolerance)
                        {
                            connectionCount++;
                            connections.Add(new PolylineConnectionInfo
                            {
                                PolylineId = objId,
                                VertexIndex = i,
                                ConnectionPoint = vertex,
                                Distance = distance,
                                IsStartPoint = (i == 0),
                                IsEndPoint = (i == pline.NumberOfVertices - 1)
                            });
                        }
                    }
                }
                else if (objId.ObjectClass == RXClass.GetClass(typeof(Polyline2d)))
                {
                    polylineCount++;
                    Polyline2d pline2d = tr.GetObject(objId, OpenMode.ForRead) as Polyline2d;
                    int vertexIndex = 0;
                    
                    // 計算總頂點數
                    int totalVertices = 0;
                    foreach (ObjectId vid in pline2d) totalVertices++;
                    
                    vertexIndex = 0;
                    foreach (ObjectId vertexId in pline2d)
                    {
                        Vertex2d vertex = tr.GetObject(vertexId, OpenMode.ForRead) as Vertex2d;
                        Point3d vertexPoint = new Point3d(vertex.Position.X, vertex.Position.Y, 0);
                        double distance = vertexPoint.DistanceTo(blockPosition);
                        
                        // 只處理端點且在容差範圍內
                        if ((vertexIndex == 0 || vertexIndex == totalVertices - 1) && distance <= tolerance)
                        {
                            connectionCount++;
                            connections.Add(new PolylineConnectionInfo
                            {
                                PolylineId = objId,
                                VertexIndex = vertexIndex,
                                ConnectionPoint = vertexPoint,
                                Distance = distance,
                                IsStartPoint = (vertexIndex == 0),
                                IsEndPoint = (vertexIndex == totalVertices - 1)
                            });
                        }
                        vertexIndex++;
                    }
                }
            }

            // 加入調試信息
            ed.WriteMessage($"\n[聚合線檢測] 檢查了 {polylineCount} 個聚合線，找到 {connectionCount} 個連接點");

            return connections;
        }

        // 全新：基於外觀交點的智能聚合線延伸/裁剪系統
        private void SmartExtendOrTrimPolylines(Transaction tr, List<PolylineConnectionInfo> connections, Point3d newBlockPosition, BlockReference newBlockRef)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            
            try
            {
                ed.WriteMessage($"\n[智能處理] 開始處理 {connections.Count} 個聚合線連接");
                
                // 獲取圖塊的實際幾何形狀
                var blockGeometry = GetBlockVisualGeometry(tr, newBlockRef);
                
                ed.WriteMessage($"\n[智能處理] 圖塊包含 {blockGeometry.Count} 個幾何實體");
                
                int processedCount = 0;
                int extendedCount = 0;
                int trimmedCount = 0;
                
                foreach (var connection in connections)
                {
                    try
                    {
                        Entity polylineEntity = tr.GetObject(connection.PolylineId, OpenMode.ForWrite) as Entity;
                        
                        // 獲取聚合線的延伸方向和長度
                        var extensionInfo = CalculatePolylineExtension(tr, connection, newBlockPosition);
                        
                        // 尋找與圖塊實體的視覺交點
                        Point3d? intersectionPoint = FindVisualIntersectionWithBlock(
                            connection.ConnectionPoint, 
                            extensionInfo.ExtensionPoint, 
                            blockGeometry);
                        
                        Point3d targetPoint;
                        bool wasTrimmed = false;
                        
                        if (intersectionPoint.HasValue)
                        {
                            // 如果找到交點，裁剪到交點
                            targetPoint = intersectionPoint.Value;
                            wasTrimmed = true;
                            trimmedCount++;
                            ed.WriteMessage($"\n[智能處理] 連接 {processedCount + 1}: 裁剪到視覺交點 ({targetPoint.X:F2}, {targetPoint.Y:F2})");
                        }
                        else
                        {
                            // 如果沒有交點，延伸到計算位置
                            targetPoint = extensionInfo.ExtensionPoint;
                            extendedCount++;
                            ed.WriteMessage($"\n[智能處理] 連接 {processedCount + 1}: 延伸到目標點 ({targetPoint.X:F2}, {targetPoint.Y:F2})");
                        }
                        
                        // 更新聚合線端點
                        bool updateSuccess = UpdatePolylineEndpoint(polylineEntity, connection, targetPoint);
                        
                        if (updateSuccess)
                        {
                            ed.WriteMessage($"\n[智能處理] 成功更新聚合線端點");
                        }
                        else
                        {
                            ed.WriteMessage($"\n[智能處理] 更新聚合線端點失敗");
                        }
                        
                        processedCount++;
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\n[錯誤] 處理連接 {processedCount + 1} 時發生錯誤: {ex.Message}");
                    }
                }
                
                ed.WriteMessage($"\n[智能處理] 完成: 總計 {processedCount} 個，延伸 {extendedCount} 個，裁剪 {trimmedCount} 個");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[錯誤] 智能處理失敗，回退到簡單延伸: {ex.Message}");
                // 如果智能處理失敗，回退到簡單的正交延伸
                FallbackToSimpleExtension(tr, connections, newBlockPosition);
            }
        }

        // 新增：獲取圖塊的視覺幾何形狀
        private List<Curve> GetBlockVisualGeometry(Transaction tr, BlockReference blockRef)
        {
            var geometry = new List<Curve>();
            
            try
            {
                // 獲取圖塊定義
                BlockTableRecord btr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                
                foreach (ObjectId entityId in btr)
                {
                    try
                    {
                        Entity entity = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
                        
                        // 只處理曲線類實體（線段、圓弧、聚合線等）
                        if (entity is Curve curve)
                        {
                            // 應用圖塊的變換矩陣
                            Curve transformedCurve = curve.GetTransformedCopy(blockRef.BlockTransform) as Curve;
                            if (transformedCurve != null)
                            {
                                geometry.Add(transformedCurve);
                            }
                        }
                        // 處理複合實體（如聚合線）
                        else if (entity is Polyline polyline)
                        {
                            Polyline transformedPline = polyline.GetTransformedCopy(blockRef.BlockTransform) as Polyline;
                            if (transformedPline != null)
                            {
                                geometry.Add(transformedPline);
                            }
                        }
                        // 處理直線
                        else if (entity is Line line)
                        {
                            Line transformedLine = line.GetTransformedCopy(blockRef.BlockTransform) as Line;
                            if (transformedLine != null)
                            {
                                geometry.Add(transformedLine);
                            }
                        }
                        // 處理圓弧
                        else if (entity is Arc arc)
                        {
                            Arc transformedArc = arc.GetTransformedCopy(blockRef.BlockTransform) as Arc;
                            if (transformedArc != null)
                            {
                                geometry.Add(transformedArc);
                            }
                        }
                        // 處理圓
                        else if (entity is Circle circle)
                        {
                            Circle transformedCircle = circle.GetTransformedCopy(blockRef.BlockTransform) as Circle;
                            if (transformedCircle != null)
                            {
                                geometry.Add(transformedCircle);
                            }
                        }
                    }
                    catch
                    {
                        // 跳過無法處理的實體
                        continue;
                    }
                }
            }
            catch
            {
                // 如果無法獲取圖塊幾何，返回空列表
            }
            
            return geometry;
        }

        // 新增：計算聚合線延伸信息
        private PolylineExtensionInfo CalculatePolylineExtension(Transaction tr, PolylineConnectionInfo connection, Point3d targetPosition)
        {
            Entity entity = tr.GetObject(connection.PolylineId, OpenMode.ForRead) as Entity;
            
            // 獲取聚合線在端點的方向
            Vector3d direction = GetPolylineDirectionAtEndpoint(tr, entity, connection);
            
            // 計算延伸距離（從連接點到目標位置的距離 + 額外緩衝）
            double baseDistance = connection.ConnectionPoint.DistanceTo(targetPosition);
            double extensionDistance = Math.Max(baseDistance, 5.0); // 最少延伸5個單位
            
            // 計算延伸點
            Point3d extensionPoint = connection.ConnectionPoint + direction * extensionDistance;
            
            return new PolylineExtensionInfo
            {
                Direction = direction,
                ExtensionDistance = extensionDistance,
                ExtensionPoint = extensionPoint
            };
        }

        // 新增：獲取聚合線在端點的方向
        private Vector3d GetPolylineDirectionAtEndpoint(Transaction tr, Entity entity, PolylineConnectionInfo connection)
        {
            Vector3d direction = Vector3d.XAxis; // 預設方向
            
            if (entity is Polyline pline)
            {
                direction = GetPolylineDirectionAtVertex(pline, connection.VertexIndex);
            }
            else if (entity is Polyline2d pline2d)
            {
                direction = GetPolyline2dDirectionAtVertex(tr, pline2d, connection.VertexIndex);
            }
            
            return direction;
        }

        // 新增：尋找與圖塊實體的視覺交點
        private Point3d? FindVisualIntersectionWithBlock(Point3d lineStart, Point3d lineEnd, List<Curve> blockGeometry)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            
            // 創建延伸線段
            using (Line extensionLine = new Line(lineStart, lineEnd))
            {
                var intersectionPoints = new List<Point3d>();
                
                foreach (Curve blockCurve in blockGeometry)
                {
                    try
                    {
                        // 計算交點
                        Point3dCollection intersections = new Point3dCollection();
                        extensionLine.IntersectWith(blockCurve, Intersect.OnBothOperands, intersections, IntPtr.Zero, IntPtr.Zero);
                        
                        // 收集所有交點
                        foreach (Point3d point in intersections)
                        {
                            intersectionPoints.Add(point);
                        }
                    }
                    catch
                    {
                        // 跳過計算失敗的曲線
                        continue;
                    }
                }
                
                ed.WriteMessage($"\n[視覺交點] 找到 {intersectionPoints.Count} 個交點");
                
                // 如果有交點，返回距離起點最近的交點
                if (intersectionPoints.Count > 0)
                {
                    Point3d nearestPoint = intersectionPoints
                        .OrderBy(p => p.DistanceTo(lineStart))
                        .First();
                    
                    ed.WriteMessage($"\n[視覺交點] 選擇最近交點: ({nearestPoint.X:F2}, {nearestPoint.Y:F2})");
                    return nearestPoint;
                }
            }
            
            return null;
        }

        // 新增：聚合線延伸信息類別
        private class PolylineExtensionInfo
        {
            public Vector3d Direction { get; set; }
            public double ExtensionDistance { get; set; }
            public Point3d ExtensionPoint { get; set; }
        }

        // 保留並改善：回退到簡單延伸邏輯
        private void FallbackToSimpleExtension(Transaction tr, List<PolylineConnectionInfo> connections, Point3d newBlockPosition)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            
            ed.WriteMessage($"\n[調試] 回退到簡單延伸模式，處理 {connections.Count} 個連接");
            
            foreach (var connection in connections)
            {
                try
                {
                    Entity entity = tr.GetObject(connection.PolylineId, OpenMode.ForWrite) as Entity;
                    
                    // 計算簡單的延伸點：直接向新圖塊位置延伸
                    Vector3d direction = (newBlockPosition - connection.ConnectionPoint).GetNormal();
                    double extensionDistance = connection.ConnectionPoint.DistanceTo(newBlockPosition) * 0.5;
                    Point3d simpleExtensionPoint = connection.ConnectionPoint + direction * extensionDistance;
                    
                    // 更新聚合線端點
                    UpdatePolylineEndpoint(entity, connection, simpleExtensionPoint);
                    
                    ed.WriteMessage($"\n[調試] 簡單延伸: ({connection.ConnectionPoint.X:F2}, {connection.ConnectionPoint.Y:F2}) -> ({simpleExtensionPoint.X:F2}, {simpleExtensionPoint.Y:F2})");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n[錯誤] 簡單延伸失敗: {ex.Message}");
                }
            }
        }

        // 保留：獲取聚合線在指定頂點的方向向量
        private Vector3d GetPolylineDirectionAtVertex(Polyline pline, int vertexIndex)
        {
            Vector3d direction = Vector3d.XAxis;

            if (vertexIndex == 0 && pline.NumberOfVertices > 1)
            {
                Point3d start = pline.GetPoint3dAt(0);
                Point3d next = pline.GetPoint3dAt(1);
                direction = (next - start).GetNormal();
            }
            else if (vertexIndex == pline.NumberOfVertices - 1 && pline.NumberOfVertices > 1)
            {
                Point3d prev = pline.GetPoint3dAt(pline.NumberOfVertices - 2);
                Point3d end = pline.GetPoint3dAt(pline.NumberOfVertices - 1);
                direction = (end - prev).GetNormal();
            }

            return direction;
        }

        // 保留：獲取2D聚合線在指定頂點的方向向量
        private Vector3d GetPolyline2dDirectionAtVertex(Transaction tr, Polyline2d pline2d, int vertexIndex)
        {
            Vector3d direction = Vector3d.XAxis;
            var vertices = new List<Point3d>();

            foreach (ObjectId vertexId in pline2d)
            {
                Vertex2d vertex = tr.GetObject(vertexId, OpenMode.ForRead) as Vertex2d;
                vertices.Add(new Point3d(vertex.Position.X, vertex.Position.Y, 0));
            }

            if (vertices.Count > 1)
            {
                if (vertexIndex == 0)
                {
                    direction = (vertices[1] - vertices[0]).GetNormal();
                }
                else if (vertexIndex == vertices.Count - 1)
                {
                    direction = (vertices[vertexIndex] - vertices[vertexIndex - 1]).GetNormal();
                }
            }

            return direction;
        }

        // 保留：更新聚合線端點的統一方法
        private bool UpdatePolylineEndpoint(Entity entity, PolylineConnectionInfo connection, Point3d newPoint)
        {
            try
            {
                if (entity is Polyline pline)
                {
                    if (connection.IsStartPoint || connection.IsEndPoint)
                    {
                        pline.SetPointAt(connection.VertexIndex, new Point2d(newPoint.X, newPoint.Y));
                        return true;
                    }
                }
                else if (entity is Polyline2d pline2d)
                {
                    int currentIndex = 0;
                    foreach (ObjectId vertexId in pline2d)
                    {
                        if (currentIndex == connection.VertexIndex && 
                            (connection.IsStartPoint || connection.IsEndPoint))
                        {
                            Vertex2d vertex = entity.Database.TransactionManager.TopTransaction.GetObject(vertexId, OpenMode.ForWrite) as Vertex2d;
                            vertex.Position = newPoint;
                            return true;
                        }
                        currentIndex++;
                    }
                }
                
                return false;
            }
            catch
            {
                return false;
            }
        }

        private void ExecuteRenameBlocks(Document doc)
        {
            Database db = doc.Database;
            var selectedLayers = GetSelectedLayers(clbRenameLayersFilter, chkFilterRenameByLayer.Checked);
            
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                foreach (DataGridViewRow row in dgvRenameBlocks.Rows)
                {
                    if (row.IsNewRow) continue;

                    string oldName = row.Cells["oldNameCol"].Value?.ToString();
                    string newName = row.Cells["newNameCol"].Value?.ToString();

                    if (string.IsNullOrEmpty(oldName) || string.IsNullOrEmpty(newName))
                        continue;

                    if (!bt.Has(oldName) || bt.Has(newName))
                        continue;

                    // 如果啟用圖層篩選，檢查是否有該圖塊在選定圖層中
                    if (chkFilterRenameByLayer.Checked && selectedLayers.Count > 0)
                    {
                        ObjectId blockId = bt[oldName];
                        BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                        
                        bool hasBlockInSelectedLayers = false;
                        foreach (ObjectId objId in modelSpace)
                        {
                            if (objId.ObjectClass == RXClass.GetClass(typeof(BlockReference)))
                            {
                                BlockReference blockRef = tr.GetObject(objId, OpenMode.ForRead) as BlockReference;
                                if (blockRef.BlockTableRecord == blockId && selectedLayers.Contains(blockRef.Layer))
                                {
                                    hasBlockInSelectedLayers = true;
                                    break;
                                }
                            }
                        }
                        
                        if (!hasBlockInSelectedLayers)
                            continue; // 如果選定圖層中沒有此圖塊的參照，則跳過
                    }

                    ObjectId blockIdToRename = bt[oldName];
                    BlockTableRecord btr = tr.GetObject(blockIdToRename, OpenMode.ForWrite) as BlockTableRecord;
                    btr.Name = newName;
                }

                tr.Commit();
            }
        }

        private void ExecuteExportBlocks(Document doc)
        {
            Database db = doc.Database;
            string fileName = txtExcelFileName.Text.Trim();
            if (string.IsNullOrEmpty(fileName))
                fileName = "圖塊清單";

            string filePath = Path.Combine(txtExcelPath.Text, fileName + ".csv");

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                var csvContent = new StringBuilder();
                csvContent.AppendLine("圖塊名稱,參照數量,旋轉角度,圖層分佈,聚合線連接,建立時間");

                foreach (ObjectId blockId in bt)
                {
                    BlockTableRecord btr = tr.GetObject(blockId, OpenMode.ForRead) as BlockTableRecord;
                    if (!btr.IsAnonymous && !btr.IsLayout)
                    {
                        int referenceCount = 0;
                        var layerDistribution = new Dictionary<string, int>();
                        var rotationAngles = new List<double>();
                        int polylineConnections = 0;

                        foreach (ObjectId objId in modelSpace)
                        {
                            if (objId.ObjectClass == RXClass.GetClass(typeof(BlockReference)))
                            {
                                BlockReference blockRef = tr.GetObject(objId, OpenMode.ForRead) as BlockReference;
                                if (blockRef.BlockTableRecord == blockId)
                                {
                                    referenceCount++;
                                    rotationAngles.Add(blockRef.Rotation * 180.0 / Math.PI); // 轉換為度數
                                    
                                    if (layerDistribution.ContainsKey(blockRef.Layer))
                                        layerDistribution[blockRef.Layer]++;
                                    else
                                        layerDistribution[blockRef.Layer] = 1;

                                    // 檢查聚合線連接
                                    var connectedPolylines = FindConnectedPolylines(tr, blockRef.Position);
                                    polylineConnections += connectedPolylines.Count;
                                }
                            }
                        }

                        string layerInfo = string.Join("; ", layerDistribution.Select(kv => $"{kv.Key}({kv.Value})"));
                        string rotationInfo = rotationAngles.Count > 0 ? 
                            $"最小:{rotationAngles.Min():F1}° 最大:{rotationAngles.Max():F1}°" : "無";
                            
                        csvContent.AppendLine($"{btr.Name},{referenceCount},\"{rotationInfo}\",\"{layerInfo}\",{polylineConnections},{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    }
                }

                tr.Commit();

                // 修正：使用UTF-8 BOM寫入CSV文件確保中文正確顯示
                var utf8WithBom = new UTF8Encoding(true);
                File.WriteAllText(filePath, csvContent.ToString(), utf8WithBom);
            }
        }
    }
}