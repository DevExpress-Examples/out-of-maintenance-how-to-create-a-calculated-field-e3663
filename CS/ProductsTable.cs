using System;
using System.Data;

namespace RichEditCalculatedField {
    public class ProductsTable : DataTable {
        public ProductsTable()
            : base() {

            this.TableName = "ProductsTable";

            this.Columns.Add("ProductID", typeof(Int32));
            this.Columns.Add("ProductName", typeof(string));
            this.Columns.Add("UnitsInStock", typeof(int));
            this.Columns.Add("UnitPrice", typeof(decimal));
            this.Constraints.Add("IDPK", this.Columns["ProductID"], true);
        }

        public static ProductsTable CreateData() {
            ProductsTable table = new ProductsTable();

            table.Rows.Add(new object[] { 1, "Chai", 39, 18 });
            table.Rows.Add(new object[] { 2, "Chang", 17, 19 });
            table.Rows.Add(new object[] { 3, "Guarana Fantastica", 20, 4.5 });
            table.Rows.Add(new object[] { 4, "Sasquatch Ale", 111, 14 });
            table.Rows.Add(new object[] { 5, "Steeleye Stout", 20, 18 });

            return table;
        }
    }
}
