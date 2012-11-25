using System;
using System.Collections.Generic;
using Devart.Data.Oracle;
using Devart.Data;
using System.Linq;
using System.Text;
using System.Data;

namespace Askme
{
    class OracleQuery
    {
        // Constructors
        public OracleQuery(string sql)
        {
            this.sql = sql;
        }

        // Public
        public void Query()
        {
            using (OracleConnection connection = new OracleConnection(
                String.Format("User Id={0}; Password={1}; Data Source={2}",
                Common.SysProp.oracleLogin, Common.SysProp.oraclePassword, Common.SysProp.oracleAliase)))
            {
                connection.Open();
                OracleDataAdapter da = new OracleDataAdapter(this.sql, connection);
                dataTable.Clear();
                da.Fill(dataTable);
                connection.Close();
            }
        }
        public DataTable DT { get { return dataTable; } }

        // Fields
        private string sql;
        private DataTable dataTable = new DataTable("oracleResult");
    }
}
