using System;

namespace Datos.cDatos
{
    /*
     Clase generada por IAR
     28/Octubre/2008 
     */
    public class DatosSQLServer :Datos.cDatos.gDatos
    {

        /// <summary>
        ///	Obtiene el objeto DatosSQLServer.
        /// </summary>
        //public DatosSQLServer()
        //{
        //}

        /// <summary>
        ///	Obtiene el objeto DatosSQLServer atraves de su cadena de conexión.
        /// </summary>
        public DatosSQLServer(string CadenaConexion)
        {
            this.CadenaConexion = CadenaConexion;
        }

        /// <summary>
        ///	Obtiene el objeto DatosSQLServer atraves de servidor y base de datos.
        /// </summary>
        public DatosSQLServer(string Servidor, string Base, string Usuario, string Pwd)
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

                        switch (this.Autenticacion)
                        {
                            case Datos.cDatos.gDatos.TipoAutenticacion.WindowsAutentication:
                                // conexion usando windows authentication
                                //System.Text.StringBuilder sCadena = new System.Text.StringBuilder("");
                                sCadena.Append("data source=<SERVIDOR>;");
                                sCadena.Append("initial catalog=<BASE>;");
                                sCadena.Append("Integrated Security=SSPI;");
                                sCadena.Append("persist security info=False;");
                                //sCadena.Append("user id=<USUARIO>;packet size=4096");
                                sCadena.Replace("<SERVIDOR>", this.Servidor);
                                sCadena.Replace("<BASE>", this.Base);

                                return sCadena.ToString();

                                break;

                            case Datos.cDatos.gDatos.TipoAutenticacion.SQLServerAutentication:
                                // conexion usando sql server authentication
                                //System.Text.StringBuilder sCadena = new System.Text.StringBuilder("");
                                sCadena.Append("data source=<SERVIDOR>;");
                                sCadena.Append("initial catalog=<BASE>;password=<PWD>;");
                                sCadena.Append("persist security info=True;");
                                sCadena.Append("user id=<USUARIO>;packet size=4096");
                                sCadena.Replace("<SERVIDOR>", this.Servidor);
                                sCadena.Replace("<BASE>", this.Base);
                                sCadena.Replace("<USUARIO>", this.Usuario);
                                sCadena.Replace("<PWD>", this.Pwd);

                                return sCadena.ToString();

                                break;
                        
                        }
                        
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
        /// Carga parámetros a usar con el stored procedure especifico
        /// </summary>
        protected override void CargarParametros(System.Data.IDbCommand Com, System.Object[] Args)
        {
            int Limite = Com.Parameters.Count;
            for (int i = 1; i < Com.Parameters.Count; i++)
            {
                System.Data.SqlClient.SqlParameter P = (System.Data.SqlClient.SqlParameter)Com.Parameters[i];
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
            System.Data.SqlClient.SqlCommand Com;
            string SQLQuery;

            SQLQuery = (string)strSQL;

            // verifica en tabla hastable por si existe el procedimiento almacenado
            if (ColComandos.Contains(SQLQuery))
                // obtiene de hashtable comando ejecutado anteriormente 
                Com = (System.Data.SqlClient.SqlCommand)ColComandos[SQLQuery];
            else
            {
                Com = new System.Data.SqlClient.SqlCommand(SQLQuery);
                Com.CommandType = System.Data.CommandType.Text;
                Com.CommandTimeout = this.mTimeOut;
                ColComandos.Add(SQLQuery, Com);
            }

            Com.Connection = (System.Data.SqlClient.SqlConnection)this.Conexion;
            Com.Transaction = (System.Data.SqlClient.SqlTransaction)this.mTransaccion;
            return (System.Data.IDbCommand)Com;
        }

        /// <summary>
        /// Crea comando para ejecutar en la base de datos. Prrocedimiento almacenado.
        /// </summary>
        protected override System.Data.IDbCommand Comando(string ProcedimientoAlmacenado)
        {
            System.Data.SqlClient.SqlCommand Com;
            // verifica en tabla hastable por si existe el procedimiento almacenado
            if (ColComandos.Contains(ProcedimientoAlmacenado))
                Com = (System.Data.SqlClient.SqlCommand)ColComandos[ProcedimientoAlmacenado];
            else
            {
                System.Data.SqlClient.SqlConnection Con2 = new System.Data.SqlClient.SqlConnection(this.CadenaConexion);
                Con2.Open();
                Com = new System.Data.SqlClient.SqlCommand(ProcedimientoAlmacenado, Con2);
                Com.CommandType = System.Data.CommandType.StoredProcedure;
                Com.CommandTimeout = this.mTimeOut;
                System.Data.SqlClient.SqlCommandBuilder.DeriveParameters(Com);
                Con2.Close();
                Con2.Dispose();
                ColComandos.Add(ProcedimientoAlmacenado, Com);

            }
            Com.Connection = (System.Data.SqlClient.SqlConnection)this.Conexion;
            Com.Transaction = (System.Data.SqlClient.SqlTransaction)this.mTransaccion;
            return (System.Data.IDbCommand)Com;
        }

        /// <summary>
        /// Crea la conexión al Servidor
        /// </summary>
        protected override System.Data.IDbConnection CrearConexion(string CadenaConexion)
        {
            return (System.Data.IDbConnection)new System.Data.SqlClient.SqlConnection(CadenaConexion);
        }

        /// <summary>
        /// Crea DataAdapter para la extracción de datos. Sentencia SQL.
        /// </summary>
        protected override System.Data.IDataAdapter CrearDataAdapter(object strSQL)
        {
            System.Data.SqlClient.SqlDataAdapter Da = new System.Data.SqlClient.SqlDataAdapter((System.Data.SqlClient.SqlCommand)Comando(strSQL));
            return (System.Data.IDataAdapter)Da;
        }

        /// <summary>
        /// Crea DataAdapter para la extracción de datos
        /// </summary>
        protected override System.Data.IDataAdapter CrearDataAdapter(string ProcedimientoAlmacenado, params System.Object[] Args)
        {
            System.Data.SqlClient.SqlDataAdapter Da = new System.Data.SqlClient.SqlDataAdapter((System.Data.SqlClient.SqlCommand)Comando(ProcedimientoAlmacenado));
            if (Args.Length != 0)
                CargarParametros(Da.SelectCommand, Args);
            return (System.Data.IDataAdapter)Da;
        }


        public override Object ObtieneListaTablas()
        {
            // Da un listado de las tablas de la base de datos seleccionada
            //////return (Object)lstHojas;
            Object obj = new Object();
            return obj;

        }

        public override Object ObtieneListaColumnas(String Ruta, String NomArchivo, String NombreTabla)
        {
            
            //////return (Object)lstHojas;
            Object obj = new Object();
            return obj;

        }

    }
}
