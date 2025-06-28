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

        public BlockChangeForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // Form properties
            this.Text = "圖塊管理工具";
            this.Size = new Size(800, 650);
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
            grpReplace.Size = new Size(740, 180);
            grpReplace.Enabled = false;

            // 匯入按鈕 - 置換圖塊
            btnImportReplace = new Button();
            btnImportReplace.Text = "匯入CSV";
            btnImportReplace.Location = new Point(10, 150);
            btnImportReplace.Size = new Size(80, 25);
            btnImportReplace.Click += BtnImportReplace_Click;

            dgvReplaceBlocks = new DataGridView();
            dgvReplaceBlocks.Location = new Point(10, 20);
            dgvReplaceBlocks.Size = new Size(720, 120);
            dgvReplaceBlocks.AllowUserToAddRows = true;
            dgvReplaceBlocks.AllowUserToDeleteRows = true;
            dgvReplaceBlocks.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            
            DataGridViewTextBoxColumn oldBlockCol = new DataGridViewTextBoxColumn();
            oldBlockCol.Name = "OldBlock";
            oldBlockCol.HeaderText = "舊圖塊名稱";
            oldBlockCol.Width = 350;
            
            DataGridViewTextBoxColumn newBlockCol = new DataGridViewTextBoxColumn();
            newBlockCol.Name = "NewBlock";
            newBlockCol.HeaderText = "新圖塊名稱";
            newBlockCol.Width = 350;
            
            dgvReplaceBlocks.Columns.Add(oldBlockCol);
            dgvReplaceBlocks.Columns.Add(newBlockCol);

            grpReplace.Controls.Add(dgvReplaceBlocks);
            grpReplace.Controls.Add(btnImportReplace);

            // Rename Blocks GroupBox
            grpRename = new GroupBox();
            grpRename.Text = "重命名圖塊設定";
            grpRename.Location = new Point(20, 240);
            grpRename.Size = new Size(740, 180);
            grpRename.Enabled = false;

            // 匯入按鈕 - 重命名圖塊
            btnImportRename = new Button();
            btnImportRename.Text = "匯入CSV";
            btnImportRename.Location = new Point(10, 150);
            btnImportRename.Size = new Size(80, 25);
            btnImportRename.Click += BtnImportRename_Click;

            dgvRenameBlocks = new DataGridView();
            dgvRenameBlocks.Location = new Point(10, 20);
            dgvRenameBlocks.Size = new Size(720, 120);
            dgvRenameBlocks.AllowUserToAddRows = true;
            dgvRenameBlocks.AllowUserToDeleteRows = true;
            dgvRenameBlocks.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            
            DataGridViewTextBoxColumn oldNameCol = new DataGridViewTextBoxColumn();
            oldNameCol.Name = "OldName";
            oldNameCol.HeaderText = "舊圖塊名稱";
            oldNameCol.Width = 350;
            
            DataGridViewTextBoxColumn newNameCol = new DataGridViewTextBoxColumn();
            newNameCol.Name = "NewName";
            newNameCol.HeaderText = "新圖塊名稱";
            newNameCol.Width = 350;
            
            dgvRenameBlocks.Columns.Add(oldNameCol);
            dgvRenameBlocks.Columns.Add(newNameCol);

            grpRename.Controls.Add(dgvRenameBlocks);
            grpRename.Controls.Add(btnImportRename);

            // Export GroupBox
            grpExport = new GroupBox();
            grpExport.Text = "輸出CSV設定";
            grpExport.Location = new Point(20, 430);
            grpExport.Size = new Size(740, 120);
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
            btnExecute.Location = new Point(600, 570);
            btnExecute.Size = new Size(80, 30);
            btnExecute.Click += BtnExecute_Click;

            btnCancel = new Button();
            btnCancel.Text = "取消";
            btnCancel.Location = new Point(690, 570);
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

        private void ChkReplaceBlocks_CheckedChanged(object sender, EventArgs e)
        {
            grpReplace.Enabled = chkReplaceBlocks.Checked;
        }

        private void ChkRenameBlocks_CheckedChanged(object sender, EventArgs e)
        {
            grpRename.Enabled = chkRenameBlocks.Checked;
        }

        private void ChkExportBlocks_CheckedChanged(object sender, EventArgs e)
        {
            grpExport.Enabled = chkExportBlocks.Checked;
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
            
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

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
                    List<ObjectId> blocksToReplace = new List<ObjectId>();
                    List<Autodesk.AutoCAD.Geometry.Matrix3d> transformations = new List<Autodesk.AutoCAD.Geometry.Matrix3d>();

                    foreach (ObjectId objId in modelSpace)
                    {
                        if (objId.ObjectClass == RXClass.GetClass(typeof(BlockReference)))
                        {
                            BlockReference blockRef = tr.GetObject(objId, OpenMode.ForRead) as BlockReference;
                            if (blockRef.Name.Equals(oldBlockName, StringComparison.OrdinalIgnoreCase))
                            {
                                blocksToReplace.Add(objId);
                                transformations.Add(blockRef.BlockTransform);
                            }
                        }
                    }

                    for (int i = 0; i < blocksToReplace.Count; i++)
                    {
                        Entity oldBlock = tr.GetObject(blocksToReplace[i], OpenMode.ForWrite) as Entity;
                        oldBlock.Erase();

                        BlockReference newBlockRef = new BlockReference(Autodesk.AutoCAD.Geometry.Point3d.Origin, newBlockId);
                        newBlockRef.TransformBy(transformations[i]);

                        modelSpace.AppendEntity(newBlockRef);
                        tr.AddNewlyCreatedDBObject(newBlockRef, true);
                    }
                }

                tr.Commit();
            }
        }

        private void ExecuteRenameBlocks(Document doc)
        {
            Database db = doc.Database;
            
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

                    ObjectId blockId = bt[oldName];
                    BlockTableRecord btr = tr.GetObject(blockId, OpenMode.ForWrite) as BlockTableRecord;
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
                csvContent.AppendLine("圖塊名稱,參照數量,建立時間");

                foreach (ObjectId blockId in bt)
                {
                    BlockTableRecord btr = tr.GetObject(blockId, OpenMode.ForRead) as BlockTableRecord;
                    if (!btr.IsAnonymous && !btr.IsLayout)
                    {
                        int referenceCount = 0;
                        foreach (ObjectId objId in modelSpace)
                        {
                            if (objId.ObjectClass == RXClass.GetClass(typeof(BlockReference)))
                            {
                                BlockReference blockRef = tr.GetObject(objId, OpenMode.ForRead) as BlockReference;
                                if (blockRef.BlockTableRecord == blockId)
                                {
                                    referenceCount++;
                                }
                            }
                        }

                        csvContent.AppendLine($"{btr.Name},{referenceCount},{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    }
                }

                tr.Commit();

                // 寫入 CSV 文件
                File.WriteAllText(filePath, csvContent.ToString(), Encoding.UTF8);
            }
        }
    }
}