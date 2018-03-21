using DocumentFormat.OpenXml.Packaging;
using Owl.Data;
using System;

namespace Owl.Excel {

    public class ExcelBuilder : IDisposable {

        private SpreadsheetDocument _package;

        public void CreateDocument(string filePath) {

            var creator = new DocumentCreator();
            this._package = creator.CreatePackage(filePath);
        }

        public void OpenDocument(string filePath) {

            throw new NotImplementedException();
        }

        public void ImportData(DataSet data) {

            DataTableImporter importer = new DataTableImporter();
            importer.ImportDataSet(this._package, data);
        }
        
        public void ImportData(DataTable table) {

            DataTableImporter importer = new DataTableImporter();
            importer.ImportDataTable(this._package, table, table.TableName);
        }

        public void ImportData(DataTable table, string sheetName) {

            DataTableImporter importer = new DataTableImporter();
            importer.ImportDataTable(this._package, table, sheetName);
        }

        public void Dispose() {

            if (this._package != null) {

                this._package.Close();
                this._package.Dispose();
            }
        }
    }
}
