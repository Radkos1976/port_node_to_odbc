using System;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Dynamic;
using System.Threading.Tasks;
using System.Linq;
using System.ComponentModel;

namespace Class_odbc
{
    public class Startup
    {        
        public async Task<object> Invoke(dynamic parameters)
        {
            try
            {
                string connectionString = (string)parameters.dsn;
                string commandString = (string)parameters.query;
                string prep = (string)parameters.prepare;
                var rek_data = new List<Prep_val>();
                List_val values = new List_val()
                {
                    Query = commandString
                };
                if (prep == "true")
                {
                    var dict = (IDictionary<string, object>)parameters.Values;
                    int i = 1;
                    while (dict.ContainsKey("Val_name" + i))
                    {
                        int len = dict.ContainsKey("Len" + i) ? (int)dict["Len" + i] : 1;
                        Prep_val tmp = new Prep_val
                        {
                            Value = dict["Value" + i],
                            Val_name = (string)dict["Val_name" + i],
                            Type = (string)dict["Type" + i],
                            Counter=0,
                            Len=len
                        };
                        rek_data.Add(tmp);
                        i++;
                    }
                    values = new List_val()
                    {
                        Query = commandString,
                        Values = rek_data
                    };
                }                
                string command = commandString.Substring(0, 6).Trim().ToLower();
                switch (command)
                {
                    case "select":
                        return await this.ExecuteQuery(connectionString, values);
                    case "insert":
                    case "update":
                    case "delete":
                        return await this.ExecuteNonQuery(connectionString, values);
                    default:
                        throw new InvalidOperationException("Unsupported type of SQL command. Only select, insert, update, delete are supported.");
                }
            }
            catch (Exception e)
            {
                throw new Exception("Parsing data error " + e, e);
            }
        }

        private async Task<object> ExecuteQuery(string connectionString,List_val prep)
        {           
            try
            {
                if (prep.Values != null)
                {
                    prep =  Create_non_parametr(prep);                    
                }
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    await connection.OpenAsync();
                    using (OleDbCommand command = new OleDbCommand(prep.Query, connection) { CommandType = CommandType.Text } )
                    {
                        if (prep.Values != null)
                        {                           
                            if (prep.Values.Count > 0)
                            {                                
                                foreach (Prep_val itm in prep.Values)
                                {                                     
                                    command.Parameters.Add(new OleDbParameter("?", GetOleDbType(itm.Type), itm.Len) { Direction = ParameterDirection.Input }).Value= Convert.ChangeType(itm.Value, itm.Value.GetType());                                                                       
                                }
                            }
                            command.Prepare();                         
                        }
                        List<object> rows = new List<object>();                        
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            IDataRecord record = (IDataRecord)reader;
                            while (await reader.ReadAsync())
                            {
                                var dataObject = new ExpandoObject() as IDictionary<string, Object>;
                                var resultRecord = new object[record.FieldCount];
                                record.GetValues(resultRecord);

                                for (int i = 0; i < record.FieldCount; i++)
                                {
                                    Type type = record.GetFieldType(i);
                                    if (resultRecord[i] is System.DBNull)
                                    {
                                        resultRecord[i] = null;
                                    }
                                    else if (type == typeof(byte[]) || type == typeof(char[]))
                                    {
                                        resultRecord[i] = Convert.ToBase64String((byte[])resultRecord[i]);
                                    }
                                    else if (type == typeof(Guid) || type == typeof(DateTime))
                                    {
                                        resultRecord[i] = resultRecord[i].ToString();
                                    }
                                    else if (type == typeof(IDataReader))
                                    {
                                        resultRecord[i] = "<IDataReader>";
                                    }

                                    dataObject.Add(record.GetName(i), resultRecord[i]);
                                }

                                rows.Add(dataObject);
                            }

                            return rows;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception("ExecuteQuery Error " + e , e);
            }
        }
        private async Task<object> ExecuteNonQuery(string connectionString, List_val prep)
        {            
            try
            {
                if (prep.Values != null)
                {
                    prep = Create_non_parametr(prep);
                }
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    
                    await connection.OpenAsync();
                    using (OleDbCommand command = new OleDbCommand(prep.Query, connection))
                    {
                        if (prep.Values != null)
                        {
                            if (prep.Values.Count > 0)
                            {
                                foreach (Prep_val itm in prep.Values)
                                {
                                    command.Parameters.Add(new OleDbParameter(itm.Val_name, GetOleDbType(itm.Type),itm.Len)).Value = Convert.ChangeType(itm.Value, itm.Value.GetType());
                                }
                            }
                            command.Prepare();
                        }
                        return await command.ExecuteNonQueryAsync();
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception("ExecuteNonQuery Error ", e);
            }
        }
        private List_val Create_non_parametr (List_val prep)
        {           
            List<Prep_val> rek_data = new List<Prep_val>();
            try
            {
                string quer = prep.Query;
                int ind = 0;
                foreach (Prep_val item in prep.Values)
                {
                    while (ind != -1)
                    {
                        ind = quer.IndexOf(item.Val_name, ind);
                        if (ind > 0)
                        {
                            Prep_val prm = new Prep_val
                            {
                                Val_name = item.Val_name,
                                Value = item.Value,
                                Type = item.Type,
                                Counter = ind,
                                Len=item.Len
                            };
                            rek_data.Add(prm);
                            ind++;
                        }
                    }
                    ind = 0;
                }
                foreach (Prep_val item in prep.Values)
                {
                    quer=quer.Replace(item.Val_name, "?");
                }
                List_val prepare = new List_val
                {
                    Values = (from m in rek_data orderby m.Counter select m).ToList(),
                    Query = quer
                };
                return prepare;
            }
            catch (Exception e)
            {
                throw new Exception("Transform parametrized Error ", e);
            }            
        }
        public class List_val 
        {
            public string Query { get; set; }
            public List<Prep_val> Values { get; set; }    
        }
        public class Prep_val
        {
            public string Val_name { get; set; }
            public dynamic Value { get; set; }
            public string Type { get; set; }
            public int Counter { get; set; }
            public int Len { get; set; }
        }

        OleDbType GetOleDbType(string typ)
        {
            try
            {                
                return (OleDbType) Enum.Parse( typeof(OleDbType),typ) ;
            }
            catch
            {
                return OleDbType.VarChar;
            }
        }       
    }
}
