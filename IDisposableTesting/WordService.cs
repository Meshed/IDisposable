using System;
using Microsoft.Office.Interop.Word;

namespace IDisposableTesting
{
    public class WordService : IDisposable
    {
        private Application _wordApplication;
        private Document _wordDocument;

        public WordService()
        {
            _wordApplication = new Application();
            _wordDocument = new Document();
        }

        public void OpenDocument()
        {
            _wordApplication.Documents.Open(@"C:\Users\markbrown\Desktop\Sandbox\SupermergeLockups_WBS.docx");
        }

        public void CloseDocument()
        {
            _wordDocument.Close(WdSaveOptions.wdDoNotSaveChanges);
            _wordDocument = null;
        }

        public void QuitApplication()
        {
            _wordApplication.Quit();
            _wordApplication = null;
        }

        public void Dispose()
        {
            if (_wordApplication != null)
            {
                if (_wordDocument != null)
                {
                    _wordDocument.Close(WdSaveOptions.wdDoNotSaveChanges);
                    _wordDocument = null;
                }

                _wordApplication.Quit();
                _wordApplication = null;
            }
        }
    }
}
