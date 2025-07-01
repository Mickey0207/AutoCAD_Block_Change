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

        // �s�W�פJ���s
        private Button btnImportReplace;
        private Button btnImportRename;
        private Button btnExportTemplate;

        // �s�W�ϼh��ܱ��
        private CheckedListBox clbReplaceLayersFilter;
        private CheckedListBox clbRenameLayersFilter;
        private CheckBox chkFilterReplaceByLayer;
        private CheckBox chkFilterRenameByLayer;
        private Button btnRefreshReplaceLayer;
        private Button btnRefreshRenameLayer;

        // �ק�G�s�W����M�E�X�u�������
        private CheckBox chkPreserveRotation;
        private CheckBox chkHandlePolylineConnection; // �ץ��G�s�W������ŧi
        
        // �s�W�G�����J���
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

        // �w�q�{�N�ưt����
        private static readonly Color PrimaryColor = Color.FromArgb(41, 128, 185);      // �D�n�Ŧ�
        private static readonly Color SecondaryColor = Color.FromArgb(52, 152, 219);    // ���n�Ŧ�  
        private static readonly Color AccentColor = Color.FromArgb(46, 204, 113);       // ���I���
        private static readonly Color DangerColor = Color.FromArgb(231, 76, 60);        // �M�I����
        private static readonly Color WarningColor = Color.FromArgb(243, 156, 18);      // ĵ�i���
        private static readonly Color LightGray = Color.FromArgb(236, 240, 241);        // �L�Ǧ�
        private static readonly Color DarkGray = Color.FromArgb(52, 73, 94);            // �`�Ǧ�
        private static readonly Color TextColor = Color.FromArgb(44, 62, 80);           // ��r�C��
        private static readonly Color CardBackground = Color.FromArgb(255, 255, 255);   // �d���I��
        private static readonly Color FormBackground = Color.FromArgb(248, 249, 250);   // ���I��

        public BlockChangeForm()
        {
            // �ץ��G��l�Ʀr�Žs�X�䴩
            InitializeEncodingSupport();
            InitializeComponent();
            ApplyModernStyling();
        }

        // �s�W�G��l�Ʀr�Žs�X�䴩
        private void InitializeEncodingSupport()
        {
            try
            {
                // ���U�s�X���Ѫ̥H�䴩����r��
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                
                // �]�w����x�s�X��UTF-8
                Console.OutputEncoding = Encoding.UTF8;
                Console.InputEncoding = Encoding.UTF8;
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"�s�X��l�ƥ���: {ex.Message}");
            }
        }

        private void ApplyModernStyling()
        {
            // �]�w���I���M�r��
            this.BackColor = FormBackground;
            
            // �ץ��G�ﵽ����r��䴩
            try
            {
                // �u���ϥ� Microsoft YaHei�A�p�G���s�b�h�ϥηL�n������Ψt�ιw�]�r��
                System.Drawing.Font baseFont = GetChineseFriendlyFont();
                this.Font = baseFont;
            }
            catch
            {
                // �p�G�r��]�w���ѡA�ϥΨt�ιw�]�r��
                this.Font = SystemFonts.DefaultFont;
            }
            
            // �]�w���D��r�˦�
            foreach (CheckBox chk in new[] { chkReplaceBlocks, chkRenameBlocks, chkExportBlocks })
            {
                StyleHeaderCheckBox(chk);
            }

            // �]�wGroupBox�˦�
            foreach (GroupBox grp in new[] { grpReplace, grpRename, grpExport })
            {
                StyleGroupBox(grp);
            }

            // �]�wDataGridView�˦�
            foreach (DataGridView dgv in new[] { dgvReplaceBlocks, dgvRenameBlocks })
            {
                StyleDataGridView(dgv);
            }

            // �]�wCheckedListBox�˦�
            foreach (CheckedListBox clb in new[] { clbReplaceLayersFilter, clbRenameLayersFilter })
            {
                StyleCheckedListBox(clb);
            }

            // �]�w���s�˦�
            StylePrimaryButton(btnExecute);
            StyleSecondaryButton(btnCancel);
            StyleAccentButton(btnImportReplace);
            StyleAccentButton(btnImportRename);
            StyleAccentButton(btnExportTemplate);
            StyleWarningButton(btnBrowsePath);
            StyleInfoButton(btnRefreshReplaceLayer);
            StyleInfoButton(btnRefreshRenameLayer);

            // �]�w��r�ؼ˦�
            foreach (TextBox txt in new[] { txtExcelFileName, txtExcelPath })
            {
                StyleTextBox(txt);
            }

            // �]�w���Ҽ˦�
            foreach (Label lbl in new[] { lblFileName, lblPath, lblRotationDegrees })
            {
                StyleLabel(lbl);
            }

            // �]�wCheckBox�˦�
            foreach (CheckBox chk in new[] { chkFilterReplaceByLayer, chkFilterRenameByLayer, 
                chkPreserveRotation, chkHandlePolylineConnection, chkApplyAdditionalRotation })
            {
                StyleCheckBox(chk);
            }

            // �]�wRadioButton�˦�
            foreach (RadioButton rb in new[] { rbRotateLeft, rbRotateRight })
            {
                StyleRadioButton(rb);
            }

            // �]�wNumericUpDown�˦�
            StyleNumericUpDown(nudRotationDegrees);
        }

        // �s�W�G�������ͦn�r��
        private System.Drawing.Font GetChineseFriendlyFont()
        {
            // ���դ��P������r��A���u������
            string[] fontNames = { 
                "Microsoft YaHei UI",  // �L�n���� UI
                "Microsoft YaHei",     // �L�n����
                "�L�n������",            // Microsoft JhengHei
                "Microsoft JhengHei UI",
                "Microsoft JhengHei",
                "SimSun",              // ����
                "SimHei",              // ����
                "NSimSun"              // �s����
            };

            foreach (string fontName in fontNames)
            {
                try
                {
                    using (var testFont = new System.Drawing.Font(fontName, 9F))
                    {
                        // �p�G�r��i�H�ЫءA��^�Ӧr��
                        return new System.Drawing.Font(fontName, 9F, FontStyle.Regular);
                    }
                }
                catch
                {
                    continue; // ���դU�@�Ӧr��
                }
            }

            // �p�G�Ҧ�����r�鳣���ѡA��^�t�ιw�]�r��
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
            
            // �Ыضꨤ��خĪG
            grp.Paint += (sender, e) =>
            {
                GroupBox gb = sender as GroupBox;
                Graphics g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                
                // ø�s�ꨤ�x�έI��
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
            
            // ���D�˦�
            dgv.ColumnHeadersDefaultCellStyle.BackColor = PrimaryColor;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgv.Font.FontFamily, 9F, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Padding = new Padding(5);
            dgv.ColumnHeadersHeight = 40;
            
            // ��ƦC�˦�
            dgv.DefaultCellStyle.BackColor = Color.White;
            dgv.DefaultCellStyle.ForeColor = TextColor;
            dgv.DefaultCellStyle.SelectionBackColor = SecondaryColor;
            dgv.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv.DefaultCellStyle.Padding = new Padding(5);
            dgv.RowTemplate.Height = 35;
            
            // ������C��
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
            
            // �Ыضꨤ��خĪG
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
                
                // ø�s���s��r
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
            chkHandlePolylineConnection = new CheckBox(); // �ץ��G�T�O��������T��l��
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
            
            // Form �򥻳]�w
            this.ClientSize = new Size(2200, 1768);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "BlockChangeForm";
            this.Padding = new Padding(20);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "�϶��޲z�u��";
            this.BackColor = FormBackground;
            
            // �D�n�\���ܰϰ� - ����
            chkReplaceBlocks.Location = new Point(30, 30);
            chkReplaceBlocks.Name = "chkReplaceBlocks";
            chkReplaceBlocks.Size = new Size(240, 44);
            chkReplaceBlocks.TabIndex = 0;
            chkReplaceBlocks.Text = "��q�m���϶�";
            chkReplaceBlocks.UseVisualStyleBackColor = true;
            chkReplaceBlocks.CheckedChanged += ChkReplaceBlocks_CheckedChanged;

            chkRenameBlocks.Location = new Point(287, 30);
            chkRenameBlocks.Name = "chkRenameBlocks";
            chkRenameBlocks.Size = new Size(275, 44);
            chkRenameBlocks.TabIndex = 1;
            chkRenameBlocks.Text = "��q���R�W�϶�";
            chkRenameBlocks.UseVisualStyleBackColor = true;
            chkRenameBlocks.CheckedChanged += ChkRenameBlocks_CheckedChanged;

            chkExportBlocks.Location = new Point(568, 28);
            chkExportBlocks.Name = "chkExportBlocks";
            chkExportBlocks.Size = new Size(238, 49);
            chkExportBlocks.TabIndex = 2;
            chkExportBlocks.Text = "��X�϶��M��";
            chkExportBlocks.UseVisualStyleBackColor = true;
            chkExportBlocks.CheckedChanged += ChkExportBlocks_CheckedChanged;

            // �m���϶��ϰ�
            grpReplace.Controls.Add(dgvReplaceBlocks);
            grpReplace.Controls.Add(btnImportReplace);
            grpReplace.Controls.Add(chkFilterReplaceByLayer);
            grpReplace.Controls.Add(btnRefreshReplaceLayer);
            grpReplace.Controls.Add(clbReplaceLayersFilter);
            grpReplace.Controls.Add(chkPreserveRotation);
            grpReplace.Controls.Add(chkHandlePolylineConnection); // �ץ��G�T�O�[�J���
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
            grpReplace.Text = "�m���϶��]�w";

            // DataGridView - �m���M��
            dgvReplaceBlocks.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvReplaceBlocks.Columns.AddRange(new DataGridViewColumn[] { oldBlockCol, newBlockCol });
            dgvReplaceBlocks.Location = new Point(20, 40);
            dgvReplaceBlocks.Name = "dgvReplaceBlocks";
            dgvReplaceBlocks.RowHeadersWidth = 102;
            dgvReplaceBlocks.Size = new Size(823, 565);
            dgvReplaceBlocks.TabIndex = 0;
            dgvReplaceBlocks.AllowUserToAddRows = true;
            dgvReplaceBlocks.AllowUserToDeleteRows = true;

            oldBlockCol.HeaderText = "�¹϶��W��";
            oldBlockCol.Name = "oldBlockCol";
            oldBlockCol.Width = 340;

            newBlockCol.HeaderText = "�s�϶��W��";
            newBlockCol.Name = "newBlockCol";
            newBlockCol.Width = 340;

            // ���s�ϰ� - �m��
            btnImportReplace.Location = new Point(1048, 491);
            btnImportReplace.Name = "btnImportReplace";
            btnImportReplace.Size = new Size(220, 125);
            btnImportReplace.TabIndex = 1;
            btnImportReplace.Text = "�פJCSV";
            btnImportReplace.UseVisualStyleBackColor = true;
            btnImportReplace.Click += BtnImportReplace_Click;

            // �ϼh�z��ϰ� - �m��
            chkFilterReplaceByLayer.Location = new Point(1274, 30);
            chkFilterReplaceByLayer.Name = "chkFilterReplaceByLayer";
            chkFilterReplaceByLayer.Size = new Size(214, 50);
            chkFilterReplaceByLayer.TabIndex = 2;
            chkFilterReplaceByLayer.Text = "���ϼh�z��";
            chkFilterReplaceByLayer.UseVisualStyleBackColor = true;
            chkFilterReplaceByLayer.CheckedChanged += ChkFilterReplaceByLayer_CheckedChanged;

            btnRefreshReplaceLayer.Location = new Point(1490, 30);
            btnRefreshReplaceLayer.Name = "btnRefreshReplaceLayer";
            btnRefreshReplaceLayer.Size = new Size(212, 60);
            btnRefreshReplaceLayer.TabIndex = 3;
            btnRefreshReplaceLayer.Text = "��s";
            btnRefreshReplaceLayer.UseVisualStyleBackColor = true;
            btnRefreshReplaceLayer.Click += BtnRefreshReplaceLayer_Click;

            clbReplaceLayersFilter.CheckOnClick = true;
            clbReplaceLayersFilter.Enabled = false;
            clbReplaceLayersFilter.Location = new Point(1274, 96);
            clbReplaceLayersFilter.Name = "clbReplaceLayersFilter";
            clbReplaceLayersFilter.Size = new Size(846, 520);
            clbReplaceLayersFilter.TabIndex = 4;

            // ���౱��ϰ�
            chkPreserveRotation.Checked = true;
            chkPreserveRotation.CheckState = CheckState.Checked;
            chkPreserveRotation.Location = new Point(855, 45);
            chkPreserveRotation.Name = "chkPreserveRotation";
            chkPreserveRotation.Size = new Size(302, 50);
            chkPreserveRotation.TabIndex = 5;
            chkPreserveRotation.Text = "�O���϶����ਤ��";
            chkPreserveRotation.UseVisualStyleBackColor = true;

            // �ץ��G�[�J�E�X�u�B�z�ﶵ
            chkHandlePolylineConnection.Location = new Point(855, 101);
            chkHandlePolylineConnection.Name = "chkHandlePolylineConnection";
            chkHandlePolylineConnection.Size = new Size(200, 30);
            chkHandlePolylineConnection.TabIndex = 6;
            chkHandlePolylineConnection.Text = "����E�X�u�B�z";
            chkHandlePolylineConnection.UseVisualStyleBackColor = true;

            chkApplyAdditionalRotation.Location = new Point(855, 137);
            chkApplyAdditionalRotation.Name = "chkApplyAdditionalRotation";
            chkApplyAdditionalRotation.Size = new Size(190, 42);
            chkApplyAdditionalRotation.TabIndex = 7;
            chkApplyAdditionalRotation.Text = "�B�~����";
            chkApplyAdditionalRotation.UseVisualStyleBackColor = true;
            chkApplyAdditionalRotation.CheckedChanged += ChkApplyAdditionalRotation_CheckedChanged;

            rbRotateLeft.Checked = true;
            rbRotateLeft.Enabled = false;
            rbRotateLeft.Location = new Point(895, 185);
            rbRotateLeft.Name = "rbRotateLeft";
            rbRotateLeft.Size = new Size(199, 48);
            rbRotateLeft.TabIndex = 8;
            rbRotateLeft.TabStop = true;
            rbRotateLeft.Text = "�V������";
            rbRotateLeft.UseVisualStyleBackColor = true;

            rbRotateRight.Enabled = false;
            rbRotateRight.Location = new Point(1100, 181);
            rbRotateRight.Name = "rbRotateRight";
            rbRotateRight.Size = new Size(190, 52);
            rbRotateRight.TabIndex = 9;
            rbRotateRight.Text = "�V�k����";
            rbRotateRight.UseVisualStyleBackColor = true;

            lblRotationDegrees.Location = new Point(941, 247);
            lblRotationDegrees.Name = "lblRotationDegrees";
            lblRotationDegrees.Size = new Size(145, 43);
            lblRotationDegrees.TabIndex = 10;
            lblRotationDegrees.Text = "����׼�:";

            nudRotationDegrees.DecimalPlaces = 1;
            nudRotationDegrees.Enabled = false;
            nudRotationDegrees.Location = new Point(1092, 247);
            nudRotationDegrees.Maximum = new decimal(new int[] { 360, 0, 0, 0 });
            nudRotationDegrees.Name = "nudRotationDegrees";
            nudRotationDegrees.Size = new Size(172, 46);
            nudRotationDegrees.TabIndex = 11;
            nudRotationDegrees.Value = new decimal(new int[] { 90, 0, 0, 0 });

            // ���R�W�϶��ϰ�
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
            grpRename.Text = "���R�W�϶��]�w";

            // DataGridView - ���R�W�M��
            dgvRenameBlocks.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvRenameBlocks.Columns.AddRange(new DataGridViewColumn[] { oldNameCol, newNameCol });
            dgvRenameBlocks.Location = new Point(20, 40);
            dgvRenameBlocks.Name = "dgvRenameBlocks";
            dgvRenameBlocks.RowHeadersWidth = 102;
            dgvRenameBlocks.Size = new Size(1022, 596);
            dgvRenameBlocks.TabIndex = 0;
            dgvRenameBlocks.AllowUserToAddRows = true;
            dgvRenameBlocks.AllowUserToDeleteRows = true;

            oldNameCol.HeaderText = "�¹϶��W��";
            oldNameCol.Name = "oldNameCol";
            oldNameCol.Width = 340;

            newNameCol.HeaderText = "�s�϶��W��";
            newNameCol.Name = "newNameCol";
            newNameCol.Width = 340;

            btnImportRename.Location = new Point(1048, 474);
            btnImportRename.Name = "btnImportRename";
            btnImportRename.Size = new Size(220, 162);
            btnImportRename.TabIndex = 1;
            btnImportRename.Text = "�פJCSV";
            btnImportRename.UseVisualStyleBackColor = true;
            btnImportRename.Click += BtnImportRename_Click;

            // �ϼh�z��ϰ� - ���R�W
            chkFilterRenameByLayer.Location = new Point(1274, 40);
            chkFilterRenameByLayer.Name = "chkFilterRenameByLayer";
            chkFilterRenameByLayer.Size = new Size(210, 58);
            chkFilterRenameByLayer.TabIndex = 2;
            chkFilterRenameByLayer.Text = "���ϼh�z��";
            chkFilterRenameByLayer.UseVisualStyleBackColor = true;
            chkFilterRenameByLayer.CheckedChanged += ChkFilterRenameByLayer_CheckedChanged;

            btnRefreshRenameLayer.Location = new Point(1490, 40);
            btnRefreshRenameLayer.Name = "btnRefreshRenameLayer";
            btnRefreshRenameLayer.Size = new Size(212, 58);
            btnRefreshRenameLayer.TabIndex = 3;
            btnRefreshRenameLayer.Text = "��s";
            btnRefreshRenameLayer.UseVisualStyleBackColor = true;
            btnRefreshRenameLayer.Click += BtnRefreshRenameLayer_Click;

            clbRenameLayersFilter.CheckOnClick = true;
            clbRenameLayersFilter.Enabled = false;
            clbRenameLayersFilter.Location = new Point(1274, 116);
            clbRenameLayersFilter.Name = "clbRenameLayersFilter";
            clbRenameLayersFilter.Size = new Size(846, 520);
            clbRenameLayersFilter.TabIndex = 4;

            // ��X�]�w�ϰ�
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
            grpExport.Text = "��XCSV�]�w";

            lblFileName.Location = new Point(16, 81);
            lblFileName.Name = "lblFileName";
            lblFileName.Size = new Size(146, 45);
            lblFileName.TabIndex = 0;
            lblFileName.Text = "�ɮצW��:";

            txtExcelFileName.Location = new Point(168, 81);
            txtExcelFileName.Name = "txtExcelFileName";
            txtExcelFileName.Size = new Size(529, 46);
            txtExcelFileName.TabIndex = 1;
            txtExcelFileName.Text = "�϶��M��";

            lblPath.Location = new Point(16, 159);
            lblPath.Name = "lblPath";
            lblPath.Size = new Size(146, 48);
            lblPath.TabIndex = 2;
            lblPath.Text = "�s����|:";

            txtExcelPath.Location = new Point(168, 156);
            txtExcelPath.Name = "txtExcelPath";
            txtExcelPath.Size = new Size(529, 46);
            txtExcelPath.TabIndex = 3;
            txtExcelPath.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            btnBrowsePath.Location = new Point(703, 154);
            btnBrowsePath.Name = "btnBrowsePath";
            btnBrowsePath.Size = new Size(147, 48);
            btnBrowsePath.TabIndex = 4;
            btnBrowsePath.Text = "�s��";
            btnBrowsePath.UseVisualStyleBackColor = true;
            btnBrowsePath.Click += BtnBrowsePath_Click;

            btnExportTemplate.Location = new Point(1055, 45);
            btnExportTemplate.Name = "btnExportTemplate";
            btnExportTemplate.Size = new Size(220, 156);
            btnExportTemplate.TabIndex = 5;
            btnExportTemplate.Text = "�ץX�d��";
            btnExportTemplate.UseVisualStyleBackColor = true;
            btnExportTemplate.Click += BtnExportTemplate_Click;

            // �D�n�ާ@���s
            btnExecute.Location = new Point(1850, 1470);
            btnExecute.Name = "btnExecute";
            btnExecute.Size = new Size(150, 50);
            btnExecute.TabIndex = 6;
            btnExecute.Text = "����ާ@";
            btnExecute.UseVisualStyleBackColor = true;
            btnExecute.Click += BtnExecute_Click;

            btnCancel.DialogResult = DialogResult.Cancel;
            btnCancel.Location = new Point(2020, 1470);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new Size(120, 50);
            btnCancel.TabIndex = 7;
            btnCancel.Text = "����";
            btnCancel.UseVisualStyleBackColor = true;

            // �[�J�Ҧ��������
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

                    // �Ƨǹϼh�W��
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
                MessageBox.Show($"���J�ϼh����: {ex.Message}", "���~", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    dialog.Filter = "CSV �ɮ�|*.csv|�Ҧ��ɮ�|*.*";
                    dialog.Title = "��ܸm���϶�CSV�ɮ�";
                    
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        ImportCsvToGrid(dialog.FileName, dgvReplaceBlocks);
                        MessageBox.Show("�m���϶��ƾڶפJ���\�I", "���\", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"�פJ����: {ex.Message}", "���~", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnImportRename_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog dialog = new OpenFileDialog())
                {
                    dialog.Filter = "CSV �ɮ�|*.csv|�Ҧ��ɮ�|*.*";
                    dialog.Title = "��ܭ��R�W�϶�CSV�ɮ�";
                    
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        ImportCsvToGrid(dialog.FileName, dgvRenameBlocks);
                        MessageBox.Show("���R�W�϶��ƾڶפJ���\�I", "���\", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"�פJ����: {ex.Message}", "���~", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnExportTemplate_Click(object sender, EventArgs e)
        {
            try
            {
                using (SaveFileDialog dialog = new SaveFileDialog())
                {
                    dialog.Filter = "CSV �ɮ�|*.csv";
                    dialog.Title = "�ץXCSV�d��";
                    dialog.FileName = "�϶��ާ@�d��.csv";
                    
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        // �Ыؽd�� - �ץ��G�ϥ�UTF-8 BOM�T�O���奿�T���
                        var templateContent = new StringBuilder();
                        templateContent.AppendLine("�¹϶��W��,�s�϶��W��");
                        templateContent.AppendLine("DOOR_OLD,DOOR_NEW");
                        templateContent.AppendLine("WINDOW_V1,WINDOW_V2");
                        templateContent.AppendLine("BLOCK1,�зǪ�");
                        
                        // �ץ��G�ϥ�UTF-8 BOM�s�X�g�J���
                        var utf8WithBom = new UTF8Encoding(true);
                        File.WriteAllText(dialog.FileName, templateContent.ToString(), utf8WithBom);
                        MessageBox.Show("�d���ץX���\�I", "���\", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"�ץX�d������: {ex.Message}", "���~", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ImportCsvToGrid(string filePath, DataGridView grid)
        {
            // �ץ��G�ϥ�UTF-8�s�XŪ��CSV���
            var lines = File.ReadAllLines(filePath, Encoding.UTF8);
            
            // �M�Ų{���ƾ�
            grid.Rows.Clear();
            
            // ���L���D��]�p�G�s�b�^
            int startIndex = 0;
            if (lines.Length > 0 && (lines[0].Contains("��") || lines[0].Contains("Old")))
            {
                startIndex = 1;
            }
            
            // �פJ�ƾ�
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
                    MessageBox.Show("�S�����ʪ� AutoCAD ���ɡC", "���~", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                MessageBox.Show("�ާ@�����I", "���\", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.DialogResult = DialogResult.OK;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"����ɵo�Ϳ��~: {ex.Message}", "���~", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                                // �ˬd�ϼh�z��
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

                                // �p�G�ҥλE�X�u�B�z�A�M��s�����E�X�u
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
                        // �R���¹϶�
                        Entity oldBlock = tr.GetObject(replaceInfo.ObjectId, OpenMode.ForWrite) as Entity;
                        oldBlock.Erase();

                        // �Ыطs�϶��ѷ�
                        BlockReference newBlockRef = new BlockReference(Point3d.Origin, newBlockId);
                        
                        Point3d finalPosition = replaceInfo.Position;

                        // �p��̲ױ��ਤ��
                        double finalRotation = 0;
                        if (chkPreserveRotation.Checked)
                        {
                            finalRotation = replaceInfo.Rotation;
                        }

                        // �s�W�G�����B�~����
                        if (chkApplyAdditionalRotation.Checked)
                        {
                            double additionalRotation = (double)nudRotationDegrees.Value * Math.PI / 180.0; // �ഫ������
                            
                            if (rbRotateLeft.Checked)
                            {
                                finalRotation += additionalRotation; // �V������]�f�ɰw�^
                            }
                            else if (rbRotateRight.Checked)
                            {
                                finalRotation -= additionalRotation; // �V�k����]���ɰw�^
                            }
                        }

                        // �����ܴ�
                        newBlockRef.TransformBy(Matrix3d.Displacement(finalPosition - Point3d.Origin));
                        newBlockRef.Rotation = finalRotation;
                        newBlockRef.ScaleFactors = replaceInfo.ScaleFactors;
                        newBlockRef.Layer = replaceInfo.Layer;

                        modelSpace.AppendEntity(newBlockRef);
                        tr.AddNewlyCreatedDBObject(newBlockRef, true);

                        // �B�z�E�X�u�s���]���ੵ��/���ŦҼ{�϶���ɡ^
                        if (chkHandlePolylineConnection.Checked && replaceInfo.ConnectedPolylines.Count > 0)
                        {
                            SmartExtendOrTrimPolylines(tr, replaceInfo.ConnectedPolylines, finalPosition, newBlockRef);
                            polylineAdjustedCount += replaceInfo.ConnectedPolylines.Count;
                        }

                        totalReplacedCount++;
                    }
                }

                tr.Commit();

                // ��ܰ��浲�G
                var resultMessage = $"���\�����F {totalReplacedCount} �ӹ϶�";
                if (chkHandlePolylineConnection.Checked && polylineAdjustedCount > 0)
                {
                    resultMessage += $"\n����B�z�F {polylineAdjustedCount} �ӻE�X�u�s��";
                }
                
                MessageBox.Show(resultMessage, "�m������", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // �x�s�϶��m���H�������O
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

        // ��s�G�E�X�u�s���H�����O
        private class PolylineConnectionInfo
        {
            public ObjectId PolylineId { get; set; }
            public int VertexIndex { get; set; }
            public Point3d ConnectionPoint { get; set; }
            public double Distance { get; set; }
            public bool IsStartPoint { get; set; }
            public bool IsEndPoint { get; set; }
        }

        // �ק�G�M��P�S�w�϶������s�����E�X�u�]�ﵽ�e�t�M�˴��޿�^
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
                    
                    // �ˬd�E�X�u�����I�O�_�b�e�t�d��
                    for (int i = 0; i < pline.NumberOfVertices; i++)
                    {
                        Point3d vertex = pline.GetPoint3dAt(i);
                        double distance = vertex.DistanceTo(blockPosition);
                        
                        // �ˬd�O�_�����I�B�b�e�t�d��
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
                    
                    // �p���`���I��
                    int totalVertices = 0;
                    foreach (ObjectId vid in pline2d) totalVertices++;
                    
                    vertexIndex = 0;
                    foreach (ObjectId vertexId in pline2d)
                    {
                        Vertex2d vertex = tr.GetObject(vertexId, OpenMode.ForRead) as Vertex2d;
                        Point3d vertexPoint = new Point3d(vertex.Position.X, vertex.Position.Y, 0);
                        double distance = vertexPoint.DistanceTo(blockPosition);
                        
                        // �u�B�z���I�B�b�e�t�d��
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

            // �[�J�ոիH��
            ed.WriteMessage($"\n[�E�X�u�˴�] �ˬd�F {polylineCount} �ӻE�X�u�A��� {connectionCount} �ӳs���I");

            return connections;
        }

        // ���s�G���~�[���I������E�X�u����/���Ũt��
        private void SmartExtendOrTrimPolylines(Transaction tr, List<PolylineConnectionInfo> connections, Point3d newBlockPosition, BlockReference newBlockRef)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            
            try
            {
                ed.WriteMessage($"\n[����B�z] �}�l�B�z {connections.Count} �ӻE�X�u�s��");
                
                // ����϶�����ڴX��Ϊ�
                var blockGeometry = GetBlockVisualGeometry(tr, newBlockRef);
                
                ed.WriteMessage($"\n[����B�z] �϶��]�t {blockGeometry.Count} �ӴX�����");
                
                int processedCount = 0;
                int extendedCount = 0;
                int trimmedCount = 0;
                
                foreach (var connection in connections)
                {
                    try
                    {
                        Entity polylineEntity = tr.GetObject(connection.PolylineId, OpenMode.ForWrite) as Entity;
                        
                        // ����E�X�u��������V�M����
                        var extensionInfo = CalculatePolylineExtension(tr, connection, newBlockPosition);
                        
                        // �M��P�϶����骺��ı���I
                        Point3d? intersectionPoint = FindVisualIntersectionWithBlock(
                            connection.ConnectionPoint, 
                            extensionInfo.ExtensionPoint, 
                            blockGeometry);
                        
                        Point3d targetPoint;
                        bool wasTrimmed = false;
                        
                        if (intersectionPoint.HasValue)
                        {
                            // �p�G�����I�A���Ũ���I
                            targetPoint = intersectionPoint.Value;
                            wasTrimmed = true;
                            trimmedCount++;
                            ed.WriteMessage($"\n[����B�z] �s�� {processedCount + 1}: ���Ũ��ı���I ({targetPoint.X:F2}, {targetPoint.Y:F2})");
                        }
                        else
                        {
                            // �p�G�S�����I�A������p���m
                            targetPoint = extensionInfo.ExtensionPoint;
                            extendedCount++;
                            ed.WriteMessage($"\n[����B�z] �s�� {processedCount + 1}: ������ؼ��I ({targetPoint.X:F2}, {targetPoint.Y:F2})");
                        }
                        
                        // ��s�E�X�u���I
                        bool updateSuccess = UpdatePolylineEndpoint(polylineEntity, connection, targetPoint);
                        
                        if (updateSuccess)
                        {
                            ed.WriteMessage($"\n[����B�z] ���\��s�E�X�u���I");
                        }
                        else
                        {
                            ed.WriteMessage($"\n[����B�z] ��s�E�X�u���I����");
                        }
                        
                        processedCount++;
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\n[���~] �B�z�s�� {processedCount + 1} �ɵo�Ϳ��~: {ex.Message}");
                    }
                }
                
                ed.WriteMessage($"\n[����B�z] ����: �`�p {processedCount} �ӡA���� {extendedCount} �ӡA���� {trimmedCount} ��");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[���~] ����B�z���ѡA�^�h��²�橵��: {ex.Message}");
                // �p�G����B�z���ѡA�^�h��²�檺���橵��
                FallbackToSimpleExtension(tr, connections, newBlockPosition);
            }
        }

        // �s�W�G����϶�����ı�X��Ϊ�
        private List<Curve> GetBlockVisualGeometry(Transaction tr, BlockReference blockRef)
        {
            var geometry = new List<Curve>();
            
            try
            {
                // ����϶��w�q
                BlockTableRecord btr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                
                foreach (ObjectId entityId in btr)
                {
                    try
                    {
                        Entity entity = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
                        
                        // �u�B�z���u������]�u�q�B�꩷�B�E�X�u���^
                        if (entity is Curve curve)
                        {
                            // ���ι϶����ܴ��x�}
                            Curve transformedCurve = curve.GetTransformedCopy(blockRef.BlockTransform) as Curve;
                            if (transformedCurve != null)
                            {
                                geometry.Add(transformedCurve);
                            }
                        }
                        // �B�z�ƦX����]�p�E�X�u�^
                        else if (entity is Polyline polyline)
                        {
                            Polyline transformedPline = polyline.GetTransformedCopy(blockRef.BlockTransform) as Polyline;
                            if (transformedPline != null)
                            {
                                geometry.Add(transformedPline);
                            }
                        }
                        // �B�z���u
                        else if (entity is Line line)
                        {
                            Line transformedLine = line.GetTransformedCopy(blockRef.BlockTransform) as Line;
                            if (transformedLine != null)
                            {
                                geometry.Add(transformedLine);
                            }
                        }
                        // �B�z�꩷
                        else if (entity is Arc arc)
                        {
                            Arc transformedArc = arc.GetTransformedCopy(blockRef.BlockTransform) as Arc;
                            if (transformedArc != null)
                            {
                                geometry.Add(transformedArc);
                            }
                        }
                        // �B�z��
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
                        // ���L�L�k�B�z������
                        continue;
                    }
                }
            }
            catch
            {
                // �p�G�L�k����϶��X��A��^�ŦC��
            }
            
            return geometry;
        }

        // �s�W�G�p��E�X�u�����H��
        private PolylineExtensionInfo CalculatePolylineExtension(Transaction tr, PolylineConnectionInfo connection, Point3d targetPosition)
        {
            Entity entity = tr.GetObject(connection.PolylineId, OpenMode.ForRead) as Entity;
            
            // ����E�X�u�b���I����V
            Vector3d direction = GetPolylineDirectionAtEndpoint(tr, entity, connection);
            
            // �p�⩵���Z���]�q�s���I��ؼЦ�m���Z�� + �B�~�w�ġ^
            double baseDistance = connection.ConnectionPoint.DistanceTo(targetPosition);
            double extensionDistance = Math.Max(baseDistance, 5.0); // �̤֩���5�ӳ��
            
            // �p�⩵���I
            Point3d extensionPoint = connection.ConnectionPoint + direction * extensionDistance;
            
            return new PolylineExtensionInfo
            {
                Direction = direction,
                ExtensionDistance = extensionDistance,
                ExtensionPoint = extensionPoint
            };
        }

        // �s�W�G����E�X�u�b���I����V
        private Vector3d GetPolylineDirectionAtEndpoint(Transaction tr, Entity entity, PolylineConnectionInfo connection)
        {
            Vector3d direction = Vector3d.XAxis; // �w�]��V
            
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

        // �s�W�G�M��P�϶����骺��ı���I
        private Point3d? FindVisualIntersectionWithBlock(Point3d lineStart, Point3d lineEnd, List<Curve> blockGeometry)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            
            // �Ыة����u�q
            using (Line extensionLine = new Line(lineStart, lineEnd))
            {
                var intersectionPoints = new List<Point3d>();
                
                foreach (Curve blockCurve in blockGeometry)
                {
                    try
                    {
                        // �p����I
                        Point3dCollection intersections = new Point3dCollection();
                        extensionLine.IntersectWith(blockCurve, Intersect.OnBothOperands, intersections, IntPtr.Zero, IntPtr.Zero);
                        
                        // �����Ҧ����I
                        foreach (Point3d point in intersections)
                        {
                            intersectionPoints.Add(point);
                        }
                    }
                    catch
                    {
                        // ���L�p�⥢�Ѫ����u
                        continue;
                    }
                }
                
                ed.WriteMessage($"\n[��ı���I] ��� {intersectionPoints.Count} �ӥ��I");
                
                // �p�G�����I�A��^�Z���_�I�̪񪺥��I
                if (intersectionPoints.Count > 0)
                {
                    Point3d nearestPoint = intersectionPoints
                        .OrderBy(p => p.DistanceTo(lineStart))
                        .First();
                    
                    ed.WriteMessage($"\n[��ı���I] ��̪ܳ���I: ({nearestPoint.X:F2}, {nearestPoint.Y:F2})");
                    return nearestPoint;
                }
            }
            
            return null;
        }

        // �s�W�G�E�X�u�����H�����O
        private class PolylineExtensionInfo
        {
            public Vector3d Direction { get; set; }
            public double ExtensionDistance { get; set; }
            public Point3d ExtensionPoint { get; set; }
        }

        // �O�d�çﵽ�G�^�h��²�橵���޿�
        private void FallbackToSimpleExtension(Transaction tr, List<PolylineConnectionInfo> connections, Point3d newBlockPosition)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            
            ed.WriteMessage($"\n[�ո�] �^�h��²�橵���Ҧ��A�B�z {connections.Count} �ӳs��");
            
            foreach (var connection in connections)
            {
                try
                {
                    Entity entity = tr.GetObject(connection.PolylineId, OpenMode.ForWrite) as Entity;
                    
                    // �p��²�檺�����I�G�����V�s�϶���m����
                    Vector3d direction = (newBlockPosition - connection.ConnectionPoint).GetNormal();
                    double extensionDistance = connection.ConnectionPoint.DistanceTo(newBlockPosition) * 0.5;
                    Point3d simpleExtensionPoint = connection.ConnectionPoint + direction * extensionDistance;
                    
                    // ��s�E�X�u���I
                    UpdatePolylineEndpoint(entity, connection, simpleExtensionPoint);
                    
                    ed.WriteMessage($"\n[�ո�] ²�橵��: ({connection.ConnectionPoint.X:F2}, {connection.ConnectionPoint.Y:F2}) -> ({simpleExtensionPoint.X:F2}, {simpleExtensionPoint.Y:F2})");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n[���~] ²�橵������: {ex.Message}");
                }
            }
        }

        // �O�d�G����E�X�u�b���w���I����V�V�q
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

        // �O�d�G���2D�E�X�u�b���w���I����V�V�q
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

        // �O�d�G��s�E�X�u���I���Τ@��k
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

                    // �p�G�ҥιϼh�z��A�ˬd�O�_���ӹ϶��b��w�ϼh��
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
                            continue; // �p�G��w�ϼh���S�����϶����ѷӡA�h���L
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
                fileName = "�϶��M��";

            string filePath = Path.Combine(txtExcelPath.Text, fileName + ".csv");

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                var csvContent = new StringBuilder();
                csvContent.AppendLine("�϶��W��,�ѷӼƶq,���ਤ��,�ϼh���G,�E�X�u�s��,�إ߮ɶ�");

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
                                    rotationAngles.Add(blockRef.Rotation * 180.0 / Math.PI); // �ഫ���׼�
                                    
                                    if (layerDistribution.ContainsKey(blockRef.Layer))
                                        layerDistribution[blockRef.Layer]++;
                                    else
                                        layerDistribution[blockRef.Layer] = 1;

                                    // �ˬd�E�X�u�s��
                                    var connectedPolylines = FindConnectedPolylines(tr, blockRef.Position);
                                    polylineConnections += connectedPolylines.Count;
                                }
                            }
                        }

                        string layerInfo = string.Join("; ", layerDistribution.Select(kv => $"{kv.Key}({kv.Value})"));
                        string rotationInfo = rotationAngles.Count > 0 ? 
                            $"�̤p:{rotationAngles.Min():F1}�X �̤j:{rotationAngles.Max():F1}�X" : "�L";
                            
                        csvContent.AppendLine($"{btr.Name},{referenceCount},\"{rotationInfo}\",\"{layerInfo}\",{polylineConnections},{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    }
                }

                tr.Commit();

                // �ץ��G�ϥ�UTF-8 BOM�g�JCSV���T�O���奿�T���
                var utf8WithBom = new UTF8Encoding(true);
                File.WriteAllText(filePath, csvContent.ToString(), utf8WithBom);
            }
        }
    }
}