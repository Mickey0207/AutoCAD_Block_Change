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

        public BlockChangeForm()
        {
            InitializeComponent();
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
            // 
            // chkReplaceBlocks
            // 
            chkReplaceBlocks.Location = new Point(22, 31);
            chkReplaceBlocks.Name = "chkReplaceBlocks";
            chkReplaceBlocks.Size = new Size(248, 39);
            chkReplaceBlocks.TabIndex = 0;
            chkReplaceBlocks.Text = "��q�m���϶�";
            chkReplaceBlocks.CheckedChanged += ChkReplaceBlocks_CheckedChanged;
            // 
            // chkRenameBlocks
            // 
            chkRenameBlocks.Location = new Point(289, 25);
            chkRenameBlocks.Name = "chkRenameBlocks";
            chkRenameBlocks.Size = new Size(293, 50);
            chkRenameBlocks.TabIndex = 1;
            chkRenameBlocks.Text = "��q���R�W�϶�";
            chkRenameBlocks.CheckedChanged += ChkRenameBlocks_CheckedChanged;
            // 
            // chkExportBlocks
            // 
            chkExportBlocks.Location = new Point(607, 27);
            chkExportBlocks.Name = "chkExportBlocks";
            chkExportBlocks.Size = new Size(323, 46);
            chkExportBlocks.TabIndex = 2;
            chkExportBlocks.Text = "��X�϶��M���CSV";
            chkExportBlocks.CheckedChanged += ChkExportBlocks_CheckedChanged;
            // 
            // grpReplace
            // 
            grpReplace.Controls.Add(dgvReplaceBlocks);
            grpReplace.Controls.Add(btnImportReplace);
            grpReplace.Controls.Add(chkFilterReplaceByLayer);
            grpReplace.Controls.Add(btnRefreshReplaceLayer);
            grpReplace.Controls.Add(clbReplaceLayersFilter);
            grpReplace.Controls.Add(chkPreserveRotation);
            grpReplace.Controls.Add(chkApplyAdditionalRotation);
            grpReplace.Controls.Add(rbRotateLeft);
            grpReplace.Controls.Add(rbRotateRight);
            grpReplace.Controls.Add(lblRotationDegrees);
            grpReplace.Controls.Add(nudRotationDegrees);
            grpReplace.Enabled = false;
            grpReplace.Location = new Point(10, 86);
            grpReplace.Name = "grpReplace";
            grpReplace.Size = new Size(2138, 544);
            grpReplace.TabIndex = 3;
            grpReplace.TabStop = false;
            grpReplace.Text = "�m���϶��]�w";
            // 
            // dgvReplaceBlocks
            // 
            dgvReplaceBlocks.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvReplaceBlocks.Columns.AddRange(new DataGridViewColumn[] { oldBlockCol, newBlockCol });
            dgvReplaceBlocks.Location = new Point(6, 45);
            dgvReplaceBlocks.Name = "dgvReplaceBlocks";
            dgvReplaceBlocks.RowHeadersWidth = 102;
            dgvReplaceBlocks.Size = new Size(660, 499);
            dgvReplaceBlocks.TabIndex = 0;
            // 
            // oldBlockCol
            // 
            oldBlockCol.HeaderText = "�¹϶��W��";
            oldBlockCol.MinimumWidth = 12;
            oldBlockCol.Name = "oldBlockCol";
            oldBlockCol.Width = 270;
            // 
            // newBlockCol
            // 
            newBlockCol.HeaderText = "�s�϶��W��";
            newBlockCol.MinimumWidth = 12;
            newBlockCol.Name = "newBlockCol";
            newBlockCol.Width = 270;
            // 
            // btnImportReplace
            // 
            btnImportReplace.Location = new Point(672, 473);
            btnImportReplace.Name = "btnImportReplace";
            btnImportReplace.Size = new Size(235, 65);
            btnImportReplace.TabIndex = 1;
            btnImportReplace.Text = "�פJCSV";
            btnImportReplace.Click += BtnImportReplace_Click;
            // 
            // chkFilterReplaceByLayer
            // 
            chkFilterReplaceByLayer.Location = new Point(1561, 20);
            chkFilterReplaceByLayer.Name = "chkFilterReplaceByLayer";
            chkFilterReplaceByLayer.Size = new Size(314, 74);
            chkFilterReplaceByLayer.TabIndex = 2;
            chkFilterReplaceByLayer.Text = "���ϼh�z��";
            chkFilterReplaceByLayer.CheckedChanged += ChkFilterReplaceByLayer_CheckedChanged;
            // 
            // btnRefreshReplaceLayer
            // 
            btnRefreshReplaceLayer.Location = new Point(1881, 20);
            btnRefreshReplaceLayer.Name = "btnRefreshReplaceLayer";
            btnRefreshReplaceLayer.Size = new Size(257, 74);
            btnRefreshReplaceLayer.TabIndex = 3;
            btnRefreshReplaceLayer.Text = "��s�ϼh";
            btnRefreshReplaceLayer.Click += BtnRefreshReplaceLayer_Click;
            // 
            // clbReplaceLayersFilter
            // 
            clbReplaceLayersFilter.CheckOnClick = true;
            clbReplaceLayersFilter.Enabled = false;
            clbReplaceLayersFilter.Location = new Point(1356, 100);
            clbReplaceLayersFilter.Name = "clbReplaceLayersFilter";
            clbReplaceLayersFilter.Size = new Size(780, 434);
            clbReplaceLayersFilter.TabIndex = 4;
            // 
            // chkPreserveRotation
            // 
            chkPreserveRotation.Checked = true;
            chkPreserveRotation.CheckState = CheckState.Checked;
            chkPreserveRotation.Location = new Point(685, 45);
            chkPreserveRotation.Name = "chkPreserveRotation";
            chkPreserveRotation.Size = new Size(301, 63);
            chkPreserveRotation.TabIndex = 5;
            chkPreserveRotation.Text = "�O���϶����ਤ��";
            // 
            // chkApplyAdditionalRotation
            // 
            chkApplyAdditionalRotation.Location = new Point(685, 114);
            chkApplyAdditionalRotation.Name = "chkApplyAdditionalRotation";
            chkApplyAdditionalRotation.Size = new Size(175, 59);
            chkApplyAdditionalRotation.TabIndex = 6;
            chkApplyAdditionalRotation.Text = "�B�~����";
            chkApplyAdditionalRotation.CheckedChanged += ChkApplyAdditionalRotation_CheckedChanged;
            // 
            // rbRotateLeft
            // 
            rbRotateLeft.Checked = true;
            rbRotateLeft.Enabled = false;
            rbRotateLeft.Location = new Point(723, 170);
            rbRotateLeft.Name = "rbRotateLeft";
            rbRotateLeft.Size = new Size(197, 60);
            rbRotateLeft.TabIndex = 7;
            rbRotateLeft.TabStop = true;
            rbRotateLeft.Text = "�V������";
            // 
            // rbRotateRight
            // 
            rbRotateRight.Enabled = false;
            rbRotateRight.Location = new Point(723, 227);
            rbRotateRight.Name = "rbRotateRight";
            rbRotateRight.Size = new Size(178, 60);
            rbRotateRight.TabIndex = 8;
            rbRotateRight.Text = "�V�k����";
            // 
            // lblRotationDegrees
            // 
            lblRotationDegrees.Location = new Point(723, 290);
            lblRotationDegrees.Name = "lblRotationDegrees";
            lblRotationDegrees.Size = new Size(149, 44);
            lblRotationDegrees.TabIndex = 9;
            lblRotationDegrees.Text = "����׼�:";
            // 
            // nudRotationDegrees
            // 
            nudRotationDegrees.DecimalPlaces = 1;
            nudRotationDegrees.Enabled = false;
            nudRotationDegrees.Location = new Point(878, 288);
            nudRotationDegrees.Maximum = new decimal(new int[] { 360, 0, 0, 0 });
            nudRotationDegrees.Name = "nudRotationDegrees";
            nudRotationDegrees.Size = new Size(141, 46);
            nudRotationDegrees.TabIndex = 10;
            nudRotationDegrees.Value = new decimal(new int[] { 90, 0, 0, 0 });
            // 
            // grpRename
            // 
            grpRename.Controls.Add(dgvRenameBlocks);
            grpRename.Controls.Add(btnImportRename);
            grpRename.Controls.Add(chkFilterRenameByLayer);
            grpRename.Controls.Add(btnRefreshRenameLayer);
            grpRename.Controls.Add(clbRenameLayersFilter);
            grpRename.Enabled = false;
            grpRename.Location = new Point(12, 675);
            grpRename.Name = "grpRename";
            grpRename.Size = new Size(2134, 534);
            grpRename.TabIndex = 4;
            grpRename.TabStop = false;
            grpRename.Text = "���R�W�϶��]�w";
            // 
            // dgvRenameBlocks
            // 
            dgvRenameBlocks.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvRenameBlocks.Columns.AddRange(new DataGridViewColumn[] { oldNameCol, newNameCol });
            dgvRenameBlocks.Location = new Point(0, 45);
            dgvRenameBlocks.Name = "dgvRenameBlocks";
            dgvRenameBlocks.RowHeadersWidth = 102;
            dgvRenameBlocks.Size = new Size(664, 515);
            dgvRenameBlocks.TabIndex = 0;
            // 
            // oldNameCol
            // 
            oldNameCol.HeaderText = "�¹϶��W��";
            oldNameCol.MinimumWidth = 12;
            oldNameCol.Name = "oldNameCol";
            oldNameCol.Width = 270;
            // 
            // newNameCol
            // 
            newNameCol.HeaderText = "�s�϶��W��";
            newNameCol.MinimumWidth = 12;
            newNameCol.Name = "newNameCol";
            newNameCol.Width = 270;
            // 
            // btnImportRename
            // 
            btnImportRename.Location = new Point(670, 465);
            btnImportRename.Name = "btnImportRename";
            btnImportRename.Size = new Size(235, 69);
            btnImportRename.TabIndex = 1;
            btnImportRename.Text = "�פJCSV";
            btnImportRename.Click += BtnImportRename_Click;
            // 
            // chkFilterRenameByLayer
            // 
            chkFilterRenameByLayer.Location = new Point(1559, 21);
            chkFilterRenameByLayer.Name = "chkFilterRenameByLayer";
            chkFilterRenameByLayer.Size = new Size(314, 72);
            chkFilterRenameByLayer.TabIndex = 2;
            chkFilterRenameByLayer.Text = "���ϼh�z��";
            chkFilterRenameByLayer.CheckedChanged += ChkFilterRenameByLayer_CheckedChanged;
            // 
            // btnRefreshRenameLayer
            // 
            btnRefreshRenameLayer.Location = new Point(1879, 20);
            btnRefreshRenameLayer.Name = "btnRefreshRenameLayer";
            btnRefreshRenameLayer.Size = new Size(240, 72);
            btnRefreshRenameLayer.TabIndex = 3;
            btnRefreshRenameLayer.Text = "��s�ϼh";
            btnRefreshRenameLayer.Click += BtnRefreshRenameLayer_Click;
            // 
            // clbRenameLayersFilter
            // 
            clbRenameLayersFilter.CheckOnClick = true;
            clbRenameLayersFilter.Enabled = false;
            clbRenameLayersFilter.Location = new Point(1354, 94);
            clbRenameLayersFilter.Name = "clbRenameLayersFilter";
            clbRenameLayersFilter.Size = new Size(774, 434);
            clbRenameLayersFilter.TabIndex = 4;
            // 
            // grpExport
            // 
            grpExport.Controls.Add(lblFileName);
            grpExport.Controls.Add(txtExcelFileName);
            grpExport.Controls.Add(lblPath);
            grpExport.Controls.Add(txtExcelPath);
            grpExport.Controls.Add(btnBrowsePath);
            grpExport.Controls.Add(btnExportTemplate);
            grpExport.Enabled = false;
            grpExport.Location = new Point(10, 1232);
            grpExport.Name = "grpExport";
            grpExport.Size = new Size(986, 267);
            grpExport.TabIndex = 5;
            grpExport.TabStop = false;
            grpExport.Text = "��XCSV�]�w";
            // 
            // lblFileName
            // 
            lblFileName.Location = new Point(6, 64);
            lblFileName.Name = "lblFileName";
            lblFileName.Size = new Size(150, 49);
            lblFileName.TabIndex = 0;
            lblFileName.Text = "�ɮצW��:";
            // 
            // txtExcelFileName
            // 
            txtExcelFileName.Location = new Point(166, 75);
            txtExcelFileName.Name = "txtExcelFileName";
            txtExcelFileName.Size = new Size(500, 46);
            txtExcelFileName.TabIndex = 1;
            txtExcelFileName.Text = "�϶��M��";
            // 
            // lblPath
            // 
            lblPath.Location = new Point(10, 130);
            lblPath.Name = "lblPath";
            lblPath.Size = new Size(150, 43);
            lblPath.TabIndex = 2;
            lblPath.Text = "�s����|:";
            // 
            // txtExcelPath
            // 
            txtExcelPath.Location = new Point(166, 127);
            txtExcelPath.Name = "txtExcelPath";
            txtExcelPath.Size = new Size(500, 46);
            txtExcelPath.TabIndex = 3;
            txtExcelPath.Text = "C:\\Users\\notel\\Desktop";
            // 
            // btnBrowsePath
            // 
            btnBrowsePath.Location = new Point(680, 130);
            btnBrowsePath.Name = "btnBrowsePath";
            btnBrowsePath.Size = new Size(180, 46);
            btnBrowsePath.TabIndex = 4;
            btnBrowsePath.Text = "�s��...";
            btnBrowsePath.Click += BtnBrowsePath_Click;
            // 
            // btnExportTemplate
            // 
            btnExportTemplate.Location = new Point(10, 198);
            btnExportTemplate.Name = "btnExportTemplate";
            btnExportTemplate.Size = new Size(199, 58);
            btnExportTemplate.TabIndex = 5;
            btnExportTemplate.Text = "�ץX�d��";
            btnExportTemplate.Click += BtnExportTemplate_Click;
            // 
            // btnExecute
            // 
            btnExecute.Location = new Point(1673, 1423);
            btnExecute.Name = "btnExecute";
            btnExecute.Size = new Size(248, 76);
            btnExecute.TabIndex = 6;
            btnExecute.Text = "����";
            btnExecute.Click += BtnExecute_Click;
            // 
            // btnCancel
            // 
            btnCancel.DialogResult = DialogResult.Cancel;
            btnCancel.Location = new Point(1927, 1423);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new Size(225, 76);
            btnCancel.TabIndex = 7;
            btnCancel.Text = "����";
            // 
            // BlockChangeForm
            // 
            ClientSize = new Size(2164, 1511);
            Controls.Add(chkReplaceBlocks);
            Controls.Add(chkRenameBlocks);
            Controls.Add(chkExportBlocks);
            Controls.Add(grpReplace);
            Controls.Add(grpRename);
            Controls.Add(grpExport);
            Controls.Add(btnExecute);
            Controls.Add(btnCancel);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "BlockChangeForm";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "�϶��޲z�u�� v3.2 - ����E�X�u�B�z";
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
                        // �Ыؽd��
                        var templateContent = new StringBuilder();
                        templateContent.AppendLine("�¹϶��W��,�s�϶��W��");
                        templateContent.AppendLine("DOOR_OLD,DOOR_NEW");
                        templateContent.AppendLine("WINDOW_V1,WINDOW_V2");
                        templateContent.AppendLine("BLOCK1,�зǪ�");
                        
                        File.WriteAllText(dialog.FileName, templateContent.ToString(), Encoding.UTF8);
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

                    string oldBlockName = row.Cells["OldBlock"].Value?.ToString();
                    string newBlockName = row.Cells["NewBlock"].Value?.ToString();

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

        // �O�d�G����϶��������ɮء]²�ƪ����@���ƥΡ^
        private Extents3d GetBlockBoundingBox(BlockReference blockRef)
        {
            try
            {
                Extents3d bounds = blockRef.GeometricExtents;
                Vector3d buffer = new Vector3d(0.05, 0.05, 0.05);
                bounds.AddPoint(bounds.MinPoint - buffer);
                bounds.AddPoint(bounds.MaxPoint + buffer);
                return bounds;
            }
            catch
            {
                Point3d center = blockRef.Position;
                double defaultSize = 1.0;
                return new Extents3d(
                    center - new Vector3d(defaultSize, defaultSize, 0),
                    center + new Vector3d(defaultSize, defaultSize, 0)
                );
            }
        }

        // �O�d�G�ˬd�I�O�_�b�϶���ɤ��]²�ƪ����@���ƥΡ^
        private bool IsPointInsideBlock(Point3d point, Extents3d blockBounds)
        {
            return point.X >= blockBounds.MinPoint.X && point.X <= blockBounds.MaxPoint.X &&
                   point.Y >= blockBounds.MinPoint.Y && point.Y <= blockBounds.MaxPoint.Y;
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

        // �O�d�G�p�⥿�橵���I�]�@���ƥΤ�k�^
        private Point3d CalculateOrthogonalExtensionPoint(Transaction tr, PolylineConnectionInfo connection, Point3d newBlockPosition)
        {
            Entity entity = tr.GetObject(connection.PolylineId, OpenMode.ForRead) as Entity;
            Point3d extensionPoint = newBlockPosition;

            if (entity is Polyline pline)
            {
                Vector3d direction = GetPolylineDirectionAtVertex(pline, connection.VertexIndex);
                extensionPoint = CalculateOrthogonalProjection(connection.ConnectionPoint, newBlockPosition, direction);
            }
            else if (entity is Polyline2d pline2d)
            {
                Vector3d direction = GetPolyline2dDirectionAtVertex(tr, pline2d, connection.VertexIndex);
                extensionPoint = CalculateOrthogonalProjection(connection.ConnectionPoint, newBlockPosition, direction);
            }

            return extensionPoint;
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

        // �O�d�G�p�⥿���v�I
        private Point3d CalculateOrthogonalProjection(Point3d linePoint, Point3d targetPoint, Vector3d lineDirection)
        {
            Vector3d perpDirection = new Vector3d(-lineDirection.Y, lineDirection.X, 0);
            Vector3d toTarget = targetPoint - linePoint;
            
            double projectionAlongLine = toTarget.DotProduct(lineDirection);
            double projectionPerpendicular = toTarget.DotProduct(perpDirection);
            
            if (Math.Abs(projectionAlongLine) >= Math.Abs(projectionPerpendicular))
            {
                return linePoint + lineDirection * projectionAlongLine;
            }
            else
            {
                return linePoint + perpDirection * projectionPerpendicular;
            }
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

                    string oldName = row.Cells["OldName"].Value?.ToString();
                    string newName = row.Cells["NewName"].Value?.ToString();

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

                // �g�J CSV ���
                File.WriteAllText(filePath, csvContent.ToString(), Encoding.UTF8);
            }
        }
    }
}