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
        private CheckBox chkHandlePolylineConnection;
        
        // 新增：旋轉輸入控制項
        private CheckBox chkApplyAdditionalRotation;
        private RadioButton rbRotateLeft;
        private RadioButton rbRotateRight;
        private NumericUpDown nudRotationDegrees;
        private Label lblRotationDegrees;

        public BlockChangeForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // Form properties
            this.Text = "圖塊管理工具 v3.2 - 智能聚合線處理";
            this.Size = new Size(1000, 850);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;

            // CheckBoxes
            chkReplaceBlocks = new CheckBox();
            chkReplaceBlocks.Text = "批量置換圖塊";
            chkReplaceBlocks.Location = new Point(20, 20);
            chkReplaceBlocks.Size = new Size(150, 20);
            chkReplaceBlocks.CheckedChanged += ChkReplaceBlocks_CheckedChanged;

            chkRenameBlocks = new CheckBox();
            chkRenameBlocks.Text = "批量重命名圖塊";
            chkRenameBlocks.Location = new Point(200, 20);
            chkRenameBlocks.Size = new Size(150, 20);
            chkRenameBlocks.CheckedChanged += ChkRenameBlocks_CheckedChanged;

            chkExportBlocks = new CheckBox();
            chkExportBlocks.Text = "輸出圖塊清單到CSV";
            chkExportBlocks.Location = new Point(380, 20);
            chkExportBlocks.Size = new Size(150, 20);
            chkExportBlocks.CheckedChanged += ChkExportBlocks_CheckedChanged;

            // Replace Blocks GroupBox
            grpReplace = new GroupBox();
            grpReplace.Text = "置換圖塊設定";
            grpReplace.Location = new Point(20, 50);
            grpReplace.Size = new Size(940, 300);
            grpReplace.Enabled = false;

            // 匯入按鈕 - 置換圖塊
            btnImportReplace = new Button();
            btnImportReplace.Text = "匯入CSV";
            btnImportReplace.Location = new Point(10, 270);
            btnImportReplace.Size = new Size(80, 25);
            btnImportReplace.Click += BtnImportReplace_Click;

            dgvReplaceBlocks = new DataGridView();
            dgvReplaceBlocks.Location = new Point(10, 20);
            dgvReplaceBlocks.Size = new Size(550, 140);
            dgvReplaceBlocks.AllowUserToAddRows = true;
            dgvReplaceBlocks.AllowUserToDeleteRows = true;
            dgvReplaceBlocks.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            
            DataGridViewTextBoxColumn oldBlockCol = new DataGridViewTextBoxColumn();
            oldBlockCol.Name = "OldBlock";
            oldBlockCol.HeaderText = "舊圖塊名稱";
            oldBlockCol.Width = 270;
            
            DataGridViewTextBoxColumn newBlockCol = new DataGridViewTextBoxColumn();
            newBlockCol.Name = "NewBlock";
            newBlockCol.HeaderText = "新圖塊名稱";
            newBlockCol.Width = 270;
            
            dgvReplaceBlocks.Columns.Add(oldBlockCol);
            dgvReplaceBlocks.Columns.Add(newBlockCol);

            // 圖層篩選控制項 - 置換
            chkFilterReplaceByLayer = new CheckBox();
            chkFilterReplaceByLayer.Text = "按圖層篩選";
            chkFilterReplaceByLayer.Location = new Point(570, 20);
            chkFilterReplaceByLayer.Size = new Size(100, 20);
            chkFilterReplaceByLayer.CheckedChanged += ChkFilterReplaceByLayer_CheckedChanged;

            btnRefreshReplaceLayer = new Button();
            btnRefreshReplaceLayer.Text = "刷新圖層";
            btnRefreshReplaceLayer.Location = new Point(680, 18);
            btnRefreshReplaceLayer.Size = new Size(80, 25);
            btnRefreshReplaceLayer.Click += BtnRefreshReplaceLayer_Click;

            clbReplaceLayersFilter = new CheckedListBox();
            clbReplaceLayersFilter.Location = new Point(570, 45);
            clbReplaceLayersFilter.Size = new Size(350, 115);
            clbReplaceLayersFilter.CheckOnClick = true;
            clbReplaceLayersFilter.Enabled = false;

            // 旋轉功能控制項
            chkPreserveRotation = new CheckBox();
            chkPreserveRotation.Text = "保持圖塊旋轉角度";
            chkPreserveRotation.Location = new Point(10, 170);
            chkPreserveRotation.Size = new Size(150, 20);
            chkPreserveRotation.Checked = true;

            // 新增：額外旋轉功能
            chkApplyAdditionalRotation = new CheckBox();
            chkApplyAdditionalRotation.Text = "額外旋轉";
            chkApplyAdditionalRotation.Location = new Point(170, 170);
            chkApplyAdditionalRotation.Size = new Size(100, 20);
            chkApplyAdditionalRotation.CheckedChanged += ChkApplyAdditionalRotation_CheckedChanged;

            rbRotateLeft = new RadioButton();
            rbRotateLeft.Text = "向左旋轉";
            rbRotateLeft.Location = new Point(280, 170);
            rbRotateLeft.Size = new Size(80, 20);
            rbRotateLeft.Checked = true;
            rbRotateLeft.Enabled = false;

            rbRotateRight = new RadioButton();
            rbRotateRight.Text = "向右旋轉";
            rbRotateRight.Location = new Point(370, 170);
            rbRotateRight.Size = new Size(80, 20);
            rbRotateRight.Enabled = false;

            lblRotationDegrees = new Label();
            lblRotationDegrees.Text = "旋轉度數:";
            lblRotationDegrees.Location = new Point(460, 172);
            lblRotationDegrees.Size = new Size(60, 20);

            nudRotationDegrees = new NumericUpDown();
            nudRotationDegrees.Location = new Point(525, 170);
            nudRotationDegrees.Size = new Size(60, 20);
            nudRotationDegrees.Minimum = 0;
            nudRotationDegrees.Maximum = 360;
            nudRotationDegrees.Value = 90;
            nudRotationDegrees.DecimalPlaces = 1;
            nudRotationDegrees.Increment = 1;
            nudRotationDegrees.Enabled = false;

            // 聚合線功能控制項（更新描述）
            chkHandlePolylineConnection = new CheckBox();
            chkHandlePolylineConnection.Text = "智能聚合線延伸/裁剪";
            chkHandlePolylineConnection.Location = new Point(10, 195);
            chkHandlePolylineConnection.Size = new Size(150, 20);

            grpReplace.Controls.AddRange(new Control[] { 
                dgvReplaceBlocks, btnImportReplace, 
                chkFilterReplaceByLayer, btnRefreshReplaceLayer, clbReplaceLayersFilter,
                chkPreserveRotation, chkApplyAdditionalRotation,
                rbRotateLeft, rbRotateRight, lblRotationDegrees, nudRotationDegrees,
                chkHandlePolylineConnection
            });

            // Rename Blocks GroupBox
            grpRename = new GroupBox();
            grpRename.Text = "重命名圖塊設定";
            grpRename.Location = new Point(20, 360);
            grpRename.Size = new Size(940, 200);
            grpRename.Enabled = false;

            // 匯入按鈕 - 重命名圖塊
            btnImportRename = new Button();
            btnImportRename.Text = "匯入CSV";
            btnImportRename.Location = new Point(10, 170);
            btnImportRename.Size = new Size(80, 25);
            btnImportRename.Click += BtnImportRename_Click;

            dgvRenameBlocks = new DataGridView();
            dgvRenameBlocks.Location = new Point(10, 20);
            dgvRenameBlocks.Size = new Size(550, 140);
            dgvRenameBlocks.AllowUserToAddRows = true;
            dgvRenameBlocks.AllowUserToDeleteRows = true;
            dgvRenameBlocks.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            
            DataGridViewTextBoxColumn oldNameCol = new DataGridViewTextBoxColumn();
            oldNameCol.Name = "OldName";
            oldNameCol.HeaderText = "舊圖塊名稱";
            oldNameCol.Width = 270;
            
            DataGridViewTextBoxColumn newNameCol = new DataGridViewTextBoxColumn();
            newNameCol.Name = "NewName";
            newNameCol.HeaderText = "新圖塊名稱";
            newNameCol.Width = 270;
            
            dgvRenameBlocks.Columns.Add(oldNameCol);
            dgvRenameBlocks.Columns.Add(newNameCol);

            // 圖層篩選控制項 - 重命名
            chkFilterRenameByLayer = new CheckBox();
            chkFilterRenameByLayer.Text = "按圖層篩選";
            chkFilterRenameByLayer.Location = new Point(570, 20);
            chkFilterRenameByLayer.Size = new Size(100, 20);
            chkFilterRenameByLayer.CheckedChanged += ChkFilterRenameByLayer_CheckedChanged;

            btnRefreshRenameLayer = new Button();
            btnRefreshRenameLayer.Text = "刷新圖層";
            btnRefreshRenameLayer.Location = new Point(680, 18);
            btnRefreshRenameLayer.Size = new Size(80, 25);
            btnRefreshRenameLayer.Click += BtnRefreshRenameLayer_Click;

            clbRenameLayersFilter = new CheckedListBox();
            clbRenameLayersFilter.Location = new Point(570, 45);
            clbRenameLayersFilter.Size = new Size(350, 115);
            clbRenameLayersFilter.CheckOnClick = true;
            clbRenameLayersFilter.Enabled = false;

            grpRename.Controls.AddRange(new Control[] { 
                dgvRenameBlocks, btnImportRename, 
                chkFilterRenameByLayer, btnRefreshRenameLayer, clbRenameLayersFilter 
            });

            // Export GroupBox
            grpExport = new GroupBox();
            grpExport.Text = "輸出CSV設定";
            grpExport.Location = new Point(20, 570);
            grpExport.Size = new Size(940, 120);
            grpExport.Enabled = false;

            Label lblFileName = new Label();
            lblFileName.Text = "檔案名稱:";
            lblFileName.Location = new Point(10, 25);
            lblFileName.Size = new Size(80, 20);

            txtExcelFileName = new TextBox();
            txtExcelFileName.Location = new Point(90, 23);
            txtExcelFileName.Size = new Size(200, 20);
            txtExcelFileName.Text = "圖塊清單";

            Label lblPath = new Label();
            lblPath.Text = "存放路徑:";
            lblPath.Location = new Point(10, 55);
            lblPath.Size = new Size(80, 20);

            txtExcelPath = new TextBox();
            txtExcelPath.Location = new Point(90, 53);
            txtExcelPath.Size = new Size(500, 20);
            txtExcelPath.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            btnBrowsePath = new Button();
            btnBrowsePath.Text = "瀏覽...";
            btnBrowsePath.Location = new Point(600, 52);
            btnBrowsePath.Size = new Size(80, 23);
            btnBrowsePath.Click += BtnBrowsePath_Click;

            // 匯出範本按鈕
            btnExportTemplate = new Button();
            btnExportTemplate.Text = "匯出範本";
            btnExportTemplate.Location = new Point(10, 85);
            btnExportTemplate.Size = new Size(100, 25);
            btnExportTemplate.Click += BtnExportTemplate_Click;

            grpExport.Controls.AddRange(new Control[] { lblFileName, txtExcelFileName, lblPath, txtExcelPath, btnBrowsePath, btnExportTemplate });

            // Buttons
            btnExecute = new Button();
            btnExecute.Text = "執行";
            btnExecute.Location = new Point(800, 710);
            btnExecute.Size = new Size(80, 30);
            btnExecute.Click += BtnExecute_Click;

            btnCancel = new Button();
            btnCancel.Text = "取消";
            btnCancel.Location = new Point(890, 710);
            btnCancel.Size = new Size(80, 30);
            btnCancel.DialogResult = DialogResult.Cancel;

            // Add controls to form
            this.Controls.AddRange(new Control[] {
                chkReplaceBlocks, chkRenameBlocks, chkExportBlocks,
                grpReplace, grpRename, grpExport,
                btnExecute, btnCancel
            });

            this.ResumeLayout(false);
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
                        // 創建範本
                        var templateContent = new StringBuilder();
                        templateContent.AppendLine("舊圖塊名稱,新圖塊名稱");
                        templateContent.AppendLine("DOOR_OLD,DOOR_NEW");
                        templateContent.AppendLine("WINDOW_V1,WINDOW_V2");
                        templateContent.AppendLine("BLOCK1,標準門");
                        
                        File.WriteAllText(dialog.FileName, templateContent.ToString(), Encoding.UTF8);
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

        // 保留：獲取圖塊的實際邊界框（簡化版本作為備用）
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

        // 保留：檢查點是否在圖塊邊界內（簡化版本作為備用）
        private bool IsPointInsideBlock(Point3d point, Extents3d blockBounds)
        {
            return point.X >= blockBounds.MinPoint.X && point.X <= blockBounds.MaxPoint.X &&
                   point.Y >= blockBounds.MinPoint.Y && point.Y <= blockBounds.MaxPoint.Y;
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

        // 保留：計算正交延伸點（作為備用方法）
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

        // 保留：計算正交投影點
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

                    string oldName = row.Cells["OldName"].Value?.ToString();
                    string newName = row.Cells["NewName"].Value?.ToString();

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

                // 寫入 CSV 文件
                File.WriteAllText(filePath, csvContent.ToString(), Encoding.UTF8);
            }
        }
    }
}