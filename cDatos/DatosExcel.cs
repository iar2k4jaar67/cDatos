using System;
using System.Collections.Generic;
//using System.Linq;
//using System.Text;

namespace Datos.cDatos
{
    /*
     Clase generada por IAR vie 27NOV2015
     Con apoyo de código de VES
     */
    public class DatosExcel : Datos.cDatos.gDatos
    {

        /// <summary>
        ///	Obtiene el objeto DatosExcel.
        /// </summary>
        //public DatosExcel()
        //{
        //}

        /// <summary>
        ///	Obtiene el objeto DatosExcel atraves de su cadena de conexión.
        /// </summary>
        public DatosExcel(string CadenaConexion)
        {
            this.CadenaConexion = CadenaConexion;
        }

        /// <summary>
        ///	Obtiene el objeto DatosExcel atraves de servidor y base de datos.
        /// </summary>
        public DatosExcel(string Servidor, string Base, string Usuario, string Pwd)
        {
            this.Base = Base;
            this.Servidor = Servidor;
            this.Usuario = Usuario;
            this.Pwd = Pwd;
        }        

        
        /// <summary>
        ///	Tiempo de respuesta del servidor.
        /// </summary>
        public override int TimeOut
        {
            get
            {
                return this.mTimeOut/7200;
                //throw new Exception("The method or operation is not implemented.");
            }
            set
            {
                this.mTimeOut = value * 7200;
                //throw new Exception("The method or operation is not implemented.");
            }
        }


        /// <summary>
        ///	Crea cadena de conexión.
        /// </summary>
        public override string CadenaConexion
	    {
	        get
	        {
	            if(this.mCadenaConexion.Length==0)
	            {
	                if(this.mBase.Length!=0 && this.mServidor.Length!=0)
	                {
                        System.Text.StringBuilder sCadena = new System.Text.StringBuilder("");

                        sCadena.Append("Provider=Microsoft.ACE.OLEDB.12.0;");
                        sCadena.Append("data source=<SERVIDOR><BASE>;");
                        sCadena.Append("Extended Properties='Excel 12.0;HDR=YES'");
                        //////sCadena.Append("persist security info=True;");
                        //////sCadena.Append("user id=<USUARIO>;packet size=4096");
                        sCadena.Replace("<SERVIDOR>", this.Servidor);
                        sCadena.Replace("<BASE>", this.Base);
                        //////sCadena.Replace("<USUARIO>", this.Usuario);
                        //////sCadena.Replace("<PWD>", this.Pwd);

                        return sCadena.ToString();

	                }
                    else
	                {
                        //return "No se puede establecer la cadena de conexión";
                        System.Exception Ex = new System.Exception("No se puede establecer la cadena de conexión");
                        throw Ex;
	                }
	            }
                // se le agregó esta línea ya que marcaba error de que no todos los caminos de salida de la función se
                // estaban utilizando
                return this.mCadenaConexion;
	        }

	        set
	        {
	            this.mCadenaConexion = value;
	        }
	    }


        /// <summary>
        /// Carga parámetros a usar con el stored procedure específico
        /// </summary>
        protected override void CargarParametros(System.Data.IDbCommand Com, System.Object[] Args)
        {
            int Limite = Com.Parameters.Count;
            for (int i = 1; i < Com.Parameters.Count; i++)
            {
                //System.Data.SqlClient.SqlParameter P = (System.Data.SqlClient.SqlParameter)Com.Parameters[i];
                System.Data.OleDb.OleDbParameter P = (System.Data.OleDb.OleDbParameter)Com.Parameters[i];
                if (i <= Args.Length)
                    P.Value = Args[i - 1];
                else
                    P.Value = null;
            }
        }


        /// <summary>
        /// Colección de comandos ejecutados
        /// </summary>
        static System.Collections.Hashtable ColComandos = new System.Collections.Hashtable();

        /// <summary>
        /// Crea comando para ejecutar en la base de datos. Sentencia SQL.
        /// </summary>
        protected override System.Data.IDbCommand Comando(object strSQL)
        {
            //////System.Data.SqlClient.SqlCommand Com;
            System.Data.OleDb.OleDbCommand Com;
            string SQLQuery;

            SQLQuery = (string)strSQL;

            // verifica en tabla hastable por si existe el procedimiento almacenado
            if (ColComandos.Contains(SQLQuery))
                // obtiene de hashtable comando ejecutado anteriormente 
                //////Com = (System.Data.SqlClient.SqlCommand)ColComandos[SQLQuery];
                Com = (System.Data.OleDb.OleDbCommand)ColComandos[SQLQuery];
            else
            {
                //////Com = new System.Data.SqlClient.SqlCommand(SQLQuery);
                Com = new System.Data.OleDb.OleDbCommand(SQLQuery);
                Com.CommandType = System.Data.CommandType.Text;
                Com.CommandTimeout = this.mTimeOut;
                ColComandos.Add(SQLQuery, Com);
            }

            //////Com.Connection = (System.Data.SqlClient.SqlConnection)this.Conexion;
            //////Com.Transaction = (System.Data.SqlClient.SqlTransaction)this.mTransaccion;
            Com.Connection = (System.Data.OleDb.OleDbConnection)this.Conexion;
            Com.Transaction = (System.Data.OleDb.OleDbTransaction)this.mTransaccion;
            return (System.Data.IDbCommand)Com;
        }

        /// <summary>
        /// Crea comando para ejecutar en la base de datos. Prrocedimiento almacenado.
        /// </summary>
        protected override System.Data.IDbCommand Comando(string ProcedimientoAlmacenado)
        {
            //////System.Data.SqlClient.SqlCommand Com;
            System.Data.OleDb.OleDbCommand Com; //
            // verifica en tabla hastable por si existe el procedimiento almacenado
            if (ColComandos.Contains(ProcedimientoAlmacenado))
                //////Com = (System.Data.SqlClient.SqlCommand)ColComandos[ProcedimientoAlmacenado];
                Com = (System.Data.OleDb.OleDbCommand)ColComandos[ProcedimientoAlmacenado];
            else
            {
                //////System.Data.SqlClient.SqlConnection Con2 = new System.Data.SqlClient.SqlConnection(this.CadenaConexion);
                System.Data.OleDb.OleDbConnection Con2 = new System.Data.OleDb.OleDbConnection(this.CadenaConexion);
                Con2.Open();
                //////Com = new System.Data.SqlClient.SqlCommand(ProcedimientoAlmacenado, Con2);
                Com = new System.Data.OleDb.OleDbCommand(ProcedimientoAlmacenado, Con2);
                Com.CommandType = System.Data.CommandType.StoredProcedure;
                Com.CommandTimeout = this.mTimeOut;
                //////System.Data.SqlClient.SqlCommandBuilder.DeriveParameters(Com);
                System.Data.OleDb.OleDbCommandBuilder.DeriveParameters(Com);
                Con2.Close();
                Con2.Dispose();
                ColComandos.Add(ProcedimientoAlmacenado, Com);

            }
            //////Com.Connection = (System.Data.SqlClient.SqlConnection)this.Conexion;
            //////Com.Transaction = (System.Data.SqlClient.SqlTransaction)this.mTransaccion;
            Com.Connection = (System.Data.OleDb.OleDbConnection)this.Conexion;
            Com.Transaction = (System.Data.OleDb.OleDbTransaction)this.mTransaccion;
            return (System.Data.IDbCommand)Com;
        }

        /// <summary>
        /// Crea la conexión al Servidor
        /// </summary>
        protected override System.Data.IDbConnection CrearConexion(string CadenaConexion)
        {
            //////return (System.Data.IDbConnection)new System.Data.SqlClient.SqlConnection(CadenaConexion);
            return (System.Data.IDbConnection)new System.Data.OleDb.OleDbConnection(CadenaConexion);
        }

        /// <summary>
        /// Crea DataAdapter para la extracción de datos. Sentencia SQL.
        /// </summary>
        protected override System.Data.IDataAdapter CrearDataAdapter(object strSQL)
        {
            //////System.Data.SqlClient.SqlDataAdapter Da = new System.Data.SqlClient.SqlDataAdapter((System.Data.SqlClient.SqlCommand)Comando(strSQL));
            System.Data.OleDb.OleDbDataAdapter Da = new System.Data.OleDb.OleDbDataAdapter((System.Data.OleDb.OleDbCommand)Comando(strSQL));
            return (System.Data.IDataAdapter)Da;
        }

        /// <summary>
        /// Crea DataAdapter para la extracción de datos
        /// </summary>
        protected override System.Data.IDataAdapter CrearDataAdapter(string ProcedimientoAlmacenado, params System.Object[] Args)
        {
            //////System.Data.SqlClient.SqlDataAdapter Da = new System.Data.SqlClient.SqlDataAdapter((System.Data.SqlClient.SqlCommand)Comando(ProcedimientoAlmacenado));
            System.Data.OleDb.OleDbDataAdapter Da = new System.Data.OleDb.OleDbDataAdapter((System.Data.OleDb.OleDbCommand)Comando(ProcedimientoAlmacenado));
            if (Args.Length != 0)
                CargarParametros(Da.SelectCommand, Args);
            return (System.Data.IDataAdapter)Da;
        }

        private System.Data.DataTable GetSchemaTable(String stringconection)
        {
            System.Data.OleDb.OleDbConnection conn2 = new System.Data.OleDb.OleDbConnection();
            conn2 = (System.Data.OleDb.OleDbConnection)this.Conexion;

            System.Data.DataTable odtSchema = new System.Data.DataTable();
            
            odtSchema = conn2.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
            conn2.Close();
            conn2.Dispose();

            return odtSchema;
            
        }

        /// <summary>
        /// Obtiene las hojas que contenga el archivo excel.
        /// </summary>
        public override Object ObtieneListaTablas()
        {
            List<String> lstHojas = new List<String>(); //{ "hoja1", "hoja2", "hoja3" };

            System.Data.DataTable odt = new System.Data.DataTable();

            odt = GetSchemaTable(this.CadenaConexion);

            foreach (System.Data.DataRow rHoja in odt.Rows)
                {
                if (rHoja["TABLE_NAME"].ToString().Contains("$"))
                    {
                    // Filtered to just sheets - they all end in '$'
                    lstHojas.Add(rHoja["TABLE_NAME"].ToString());
                    }
                }
            
            // en este link existe código para obtener nombres de columnas de tablas
            // http://csharp.net-informations.com/dataset/dataset-column-definition-oledb.htm

            return (Object)lstHojas;

        }

        /// <summary>
        /// Obtiene la lista de columnas/tipo de dato de la hoja del archivo excel especificado.
        /// </summary>
        public override Object ObtieneListaColumnas(String Ruta, String NomArchivo, String NombreTabla)
        {
            
            System.Data.DataSet ods = new System.Data.DataSet();
            System.Data.DataTable odt = new System.Data.DataTable();

            this.Servidor = Ruta;
            this.Base = NomArchivo;


            System.Data.OleDb.OleDbConnection conn2 = new System.Data.OleDb.OleDbConnection();
            conn2 = (System.Data.OleDb.OleDbConnection)CrearConexion(this.CadenaConexion);

            Object strSQl = "SELECT * FROM [" + NombreTabla + "]"; // para excel se incluye caracter $ al final del nombre de la hoja (ej "hoja1$")

            CrearDataAdapter(strSQl).Fill(ods);

            odt = ods.Tables[0];

            List<string> lstCols = new List<string>();

            // link de consulta
            // IAR mier 221NOV2018
            // https://www.dotnetperls.com/dictionary
            //Dictionary<string, int> dictionary = new Dictionary<string, int>();

            //dictionary.Add("cat", 2);
            //dictionary.Add("dog", 1);
            //dictionary.Add("llama", 0);
            //dictionary.Add("iguana", -1);

            foreach (System.Data.DataColumn colTbl in odt.Columns)
            {
                
                //dictionary.Add(colTbl.ColumnName.ToString, colTbl.DataType.FullName.ToString);
                string item = colTbl.ColumnName.ToString() + ", " + colTbl.DataType.FullName.ToString();
                
                lstCols.Add(item);

            }

            // en este link existe código para obtener nombres de columnas de tablas
            // http://csharp.net-informations.com/dataset/dataset-column-definition-oledb.htm


            return (Object)lstCols;

        }

    }

}
