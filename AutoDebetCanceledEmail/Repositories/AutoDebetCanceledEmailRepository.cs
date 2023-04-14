using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoDebetCanceledEmail.Repositories
{
    public class AutoDebetCanceledEmailRepository : Base
    {
        public DataTable Get()
        {
            var data = new DataTable();
            using (SqlConnection con = new SqlConnection(conn))
            {
                try
                {
                    using (SqlCommand cmd = new SqlCommand("spAutoDebetCancelEmail"))
                    {
                        cmd.Connection = con;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 0;
                        con.Open();
                        data.Load(cmd.ExecuteReader());

                        return data;
                    }
                }
                catch (Exception ex)
                {
                    return null;
                }
                finally
                {
                    con.Close();
                }
            }
        }
    }
}
