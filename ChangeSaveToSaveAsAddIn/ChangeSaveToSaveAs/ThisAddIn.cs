using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace ChangeSaveToSaveAs
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Word.Application wordApp = this.Application;
            wordApp.DocumentBeforeSave += WordApp_DocumentBeforeSave;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Word.Application wordApp = this.Application;
            wordApp.DocumentBeforeSave -= WordApp_DocumentBeforeSave;
        }

        private void WordApp_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            //显示保存框
            this.Application.Dialogs[WdWordDialog.wdDialogFileSaveAs].Show();
            //取消保存
            Cancel = true;
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
