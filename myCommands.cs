using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;
using System.Windows.Forms;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCAD_Block_Change.MyCommands))]
[assembly: ExtensionApplication(typeof(AutoCAD_Block_Change.PluginExtension))]

namespace AutoCAD_Block_Change
{
    public class MyCommands
    {
        // 統一的圖塊管理工具指令
        [CommandMethod("BLOCKTOOL", CommandFlags.Modal)]
        public void BlockTool()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                ed.WriteMessage("\n啟動圖塊工具...");
                
                // 顯示統一的用戶界面
                using (BlockChangeForm form = new BlockChangeForm())
                {
                    DialogResult result = AcadApp.ShowModalDialog(form);
                    
                    if (result == DialogResult.OK)
                    {
                        ed.WriteMessage("\n圖塊工具操作完成。");
                    }
                    else
                    {
                        ed.WriteMessage("\n操作已取消。");
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n錯誤: {ex.Message}");
                ed.WriteMessage($"\n堆疊追蹤: {ex.StackTrace}");
            }
        }

        // 備用的中文命令
        [CommandMethod("圖塊工具", CommandFlags.Modal)]
        public void BlockToolChinese()
        {
            BlockTool(); // 調用主要方法
        }

        // 測試用的簡單命令
        [CommandMethod("BT", CommandFlags.Modal)]
        public void BlockToolSimple()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            ed.WriteMessage("\n簡化命令 BT 正常工作！");
            BlockTool(); // 調用主要方法
        }

        // 保留原有的單獨功能指令以向後兼容

        // 批量置換圖塊指令
        [CommandMethod("BLOCKCHANGE", "REPLACEBLOCKS", "置換圖塊", CommandFlags.Modal)]
        public void ReplaceBlocks()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            try
            {
                // 提示用戶輸入舊圖塊名稱
                PromptStringOptions psoOldName = new PromptStringOptions("\n請輸入要替換的舊圖塊名稱: ");
                psoOldName.AllowSpaces = false;
                PromptResult prOldName = ed.GetString(psoOldName);
                if (prOldName.Status != PromptStatus.OK)
                    return;

                string oldBlockName = prOldName.StringResult;

                // 提示用戶輸入新圖塊名稱
                PromptStringOptions psoNewName = new PromptStringOptions("\n請輸入新圖塊名稱: ");
                psoNewName.AllowSpaces = false;
                PromptResult prNewName = ed.GetString(psoNewName);
                if (prNewName.Status != PromptStatus.OK)
                    return;

                string newBlockName = prNewName.StringResult;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    
                    // 檢查舊圖塊是否存在
                    if (!bt.Has(oldBlockName))
                    {
                        ed.WriteMessage($"\n錯誤: 找不到名為 '{oldBlockName}' 的圖塊。");
                        return;
                    }

                    // 檢查新圖塊是否存在
                    if (!bt.Has(newBlockName))
                    {
                        ed.WriteMessage($"\n錯誤: 找不到名為 '{newBlockName}' 的圖塊。");
                        return;
                    }

                    // 獲取新圖塊的定義
                    ObjectId newBlockId = bt[newBlockName];

                    // 遍歷模型空間查找舊圖塊的參照
                    BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    
                    List<ObjectId> blocksToReplace = new List<ObjectId>();
                    List<Matrix3d> transformations = new List<Matrix3d>();

                    foreach (ObjectId objId in modelSpace)
                    {
                        if (objId.ObjectClass == RXClass.GetClass(typeof(BlockReference)))
                        {
                            BlockReference blockRef = tr.GetObject(objId, OpenMode.ForRead) as BlockReference;
                            if (blockRef.Name.Equals(oldBlockName, System.StringComparison.OrdinalIgnoreCase))
                            {
                                blocksToReplace.Add(objId);
                                transformations.Add(blockRef.BlockTransform);
                            }
                        }
                    }

                    // 刪除舊圖塊並插入新圖塊
                    int replacedCount = 0;
                    for (int i = 0; i < blocksToReplace.Count; i++)
                    {
                        // 刪除舊圖塊
                        Entity oldBlock = tr.GetObject(blocksToReplace[i], OpenMode.ForWrite) as Entity;
                        oldBlock.Erase();

                        // 創建新圖塊參照
                        BlockReference newBlockRef = new BlockReference(Point3d.Origin, newBlockId);
                        newBlockRef.TransformBy(transformations[i]);

                        // 添加到模型空間
                        modelSpace.AppendEntity(newBlockRef);
                        tr.AddNewlyCreatedDBObject(newBlockRef, true);

                        replacedCount++;
                    }

                    tr.Commit();
                    ed.WriteMessage($"\n成功替換了 {replacedCount} 個圖塊，從 '{oldBlockName}' 替換為 '{newBlockName}'。");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n錯誤: {ex.Message}");
            }
        }

        // 批量重命名圖塊指令
        [CommandMethod("BLOCKCHANGE", "RENAMEBLOCKS", "重命名圖塊", CommandFlags.Modal)]
        public void RenameBlocks()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            try
            {
                // 提示用戶輸入舊圖塊名稱
                PromptStringOptions psoOldName = new PromptStringOptions("\n請輸入要重命名的圖塊名稱: ");
                psoOldName.AllowSpaces = false;
                PromptResult prOldName = ed.GetString(psoOldName);
                if (prOldName.Status != PromptStatus.OK)
                    return;

                string oldBlockName = prOldName.StringResult;

                // 提示用戶輸入新圖塊名稱
                PromptStringOptions psoNewName = new PromptStringOptions("\n請輸入新的圖塊名稱: ");
                psoNewName.AllowSpaces = false;
                PromptResult prNewName = ed.GetString(psoNewName);
                if (prNewName.Status != PromptStatus.OK)
                    return;

                string newBlockName = prNewName.StringResult;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    
                    // 檢查舊圖塊是否存在
                    if (!bt.Has(oldBlockName))
                    {
                        ed.WriteMessage($"\n錯誤: 找不到名為 '{oldBlockName}' 的圖塊。");
                        return;
                    }

                    // 檢查新名稱是否已經存在
                    if (bt.Has(newBlockName))
                    {
                        ed.WriteMessage($"\n錯誤: 名為 '{newBlockName}' 的圖塊已經存在。");
                        return;
                    }

                    // 獲取圖塊定義
                    ObjectId blockId = bt[oldBlockName];
                    BlockTableRecord btr = tr.GetObject(blockId, OpenMode.ForWrite) as BlockTableRecord;
                    
                    // 重命名圖塊
                    btr.Name = newBlockName;

                    // 統計有多少個參照
                    int referenceCount = 0;
                    BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                    
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

                    tr.Commit();
                    ed.WriteMessage($"\n成功將圖塊 '{oldBlockName}' 重命名為 '{newBlockName}'。");
                    ed.WriteMessage($"\n影響了 {referenceCount} 個圖塊參照。");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n錯誤: {ex.Message}");
            }
        }

        // 列出所有圖塊名稱的輔助指令
        [CommandMethod("BLOCKCHANGE", "LISTBLOCKS", "列出圖塊", CommandFlags.Modal)]
        public void ListBlocks()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    
                    ed.WriteMessage("\n當前圖紙中的所有圖塊:");
                    ed.WriteMessage("\n=====================================");
                    
                    int count = 0;
                    foreach (ObjectId blockId in bt)
                    {
                        BlockTableRecord btr = tr.GetObject(blockId, OpenMode.ForRead) as BlockTableRecord;
                        // 只列出非匿名、非佈局的圖塊
                        if (!btr.IsAnonymous && !btr.IsLayout)
                        {
                            count++;
                            ed.WriteMessage($"\n{count}. {btr.Name}");
                        }
                    }
                    
                    if (count == 0)
                    {
                        ed.WriteMessage("\n沒有找到用戶定義的圖塊。");
                    }
                    else
                    {
                        ed.WriteMessage($"\n\n總共找到 {count} 個圖塊。");
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n錯誤: {ex.Message}");
            }
        }

        // 診斷命令
        [CommandMethod("TESTLOAD", CommandFlags.Modal)]
        public void TestLoad()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            
            ed.WriteMessage("\n=== 插件載入診斷 ===");
            ed.WriteMessage($"\n當前文檔: {(doc != null ? "已載入" : "未載入")}");
            ed.WriteMessage("\n可用命令:");
            ed.WriteMessage("\n- BLOCKTOOL");
            ed.WriteMessage("\n- 圖塊工具");
            ed.WriteMessage("\n- BT");
            ed.WriteMessage("\n- REPLACEBLOCKS");
            ed.WriteMessage("\n- RENAMEBLOCKS");
            ed.WriteMessage("\n- LISTBLOCKS");
            ed.WriteMessage("\n================");
        }

        // The CommandMethod attribute can be applied to any public  member 
        // function of any public class.
        // The function should take no arguments and return nothing.
        // If the method is an intance member then the enclosing class is 
        // intantiated for each document. If the member is a static member then
        // the enclosing class is NOT intantiated.
        //
        // NOTE: CommandMethod has overloads where you can provide helpid and
        // context menu.

        // Modal Command with localized name
        [CommandMethod("MyGroup", "MyCommand", "MyCommandLocal", CommandFlags.Modal)]
        public void MyCommand() // This method can have any name
        {
            // Put your command code here
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor ed;
            if (doc != null)
            {
                ed = doc.Editor;
                ed.WriteMessage("Hello, this is your first command.");

            }
        }

        // Modal Command with pickfirst selection
        [CommandMethod("MyGroup", "MyPickFirst", "MyPickFirstLocal", CommandFlags.Modal | CommandFlags.UsePickSet)]
        public void MyPickFirst() // This method can have any name
        {
            PromptSelectionResult result = AcadApp.DocumentManager.MdiActiveDocument.Editor.GetSelection();
            if (result.Status == PromptStatus.OK)
            {
                // There are selected entities
                // Put your command using pickfirst set code here
            }
            else
            {
                // There are no selected entities
                // Put your command code here
            }
        }

        // Application Session Command with localized name
        [CommandMethod("MyGroup", "MySessionCmd", "MySessionCmdLocal", CommandFlags.Modal | CommandFlags.Session)]
        public void MySessionCmd() // This method can have any name
        {
            // Put your command code here
        }

        // LispFunction is similar to CommandMethod but it creates a lisp 
        // callable function. Many return types are supported not just string
        // or integer.
        [LispFunction("MyLispFunction", "MyLispFunctionLocal")]
        public int MyLispFunction(ResultBuffer args) // This method can have any name
        {
            // Put your command code here

            // Return a value to the AutoCAD Lisp Interpreter
            return 1;
        }
    }
}
