using System;

namespace Datos.cDatos
{
    /// <summary>
    /// Clase abstracta de acceso a Datos.
    /// </summary>
    public abstract class gDatos
    {

        #region " Declaración de Variables "

        protected string mServidor = "";
        protected string mBase = "";
        protected string mUsuario = "";
        protected string mPwd = "";
        protected string mCadenaConexion = "";
        protected System.Data.IDbConnection mConexion;
        protected int mTimeOut;

        public enum TipoAutenticacion { WindowsAutentication = 1, SQLServerAutentication };

        // indica el tipo de autenticación
        // 1 = Windows
        // 2 = SQL Server
        private TipoAutenticacion mAutentication; 

        #endregion

        #region " Constructores "

        /// <summary>
        /// Constructores de la clase Datos.
        /// </summary>
        public gDatos()
        {
            //
            // TODO: agregar aquí la lógica del constructor
            //
        }
        #endregion

        #region " Propiedades "

        /// <summary>
        ///	Nombre del equipo servidor de datos.
        /// </summary>
        public string Servidor
        {
            get
            {
                return mServidor;
            }
            set
            {
                mServidor = value;
            }
        }

        /// <summary>
        /// Nombre de la base de datos a utilizar.
        /// </summary>
        public string Base
        {
            get
            {
                return mBase;
            }
            set
            {
                mBase = value;
            }
        }

        /// <summary>
        ///	Nombre del usuario.
        /// </summary>
        public string Usuario
        {
            get
            {
                return mUsuario;
            }
            set
            {
                mUsuario = value;
            }
        }

        /// <summary>
        /// Password del usuario.
        /// </summary>
        public string Pwd
        {
            get
            {
                return mPwd;
            }
            set
            {
                mPwd = value;
            }
        }

        /// <summary>
        /// Password del usuario.
        /// </summary>
        public TipoAutenticacion Autenticacion
        {
            get
            {
                return mAutentication;
            }
            set
            {
                mAutentication = value;
            }
        }


        /// <summary>
        /// Cadena de conexión completa a la base.
        /// </summary>
        public abstract string CadenaConexion
        {
            get;
            set;
        }

        /// <summary>
        ///	Tiempo de respuesta del Servidor.
        /// </summary>
        public abstract int TimeOut
        {
            get;
            set;
        }

        #endregion

        #region " Privadas "
        
        /// <summary>
        /// Crea u obtiene un objeto para conectarse a la base de datos.
        /// </summary>
        protected  System.Data.IDbConnection Conexion
        {
            get
            {
                if (null == mConexion)
                {
                    mConexion = CrearConexion(this.CadenaConexion);
                }
                if (mConexion.State != System.Data.ConnectionState.Open)
                    //try
                    //{

                    //}
                    //catch (Datos.cDatos ex)
                    //{
                        
                    //    throw  new MisExcepciones("Error a obtener los productos. " + ex.Message);

                    //}
                    {
                        if (mConexion.ConnectionString == "")
                        {
                            mConexion = CrearConexion(this.CadenaConexion);
                            
                        }

                        mConexion.Open();

                    }
                    

                return mConexion;
            }
        }

        #endregion

        #region " Lecturas "

        /// <summary>
        /// Obtiene un DataSet a partir de una sentencia SQL.
        /// </summary>
        public System.Data.DataSet TraerDataset(object strSQL)
        {
            System.Data.DataSet mDataSet = new System.Data.DataSet();

            try
            {
                this.CrearDataAdapter(strSQL).Fill(mDataSet);

                return mDataSet;
            }
            catch (Exception)
            {
                //return null;
                throw ;
            }
           
        }

        /// <summary>
        /// Obtiene un DataSet a partir de un Procedimiento Almacenado.
        /// </summary>
        public System.Data.DataSet TraerDataset(string ProcedimientoAlmacenado)
        {
            System.Data.DataSet mDataSet = new System.Data.DataSet();
            this.CrearDataAdapter(ProcedimientoAlmacenado).Fill(mDataSet);
            return mDataSet;

        }

        /// <summary>
        /// Obtiene un DataSet a partir de un Procedimiento Almacenado y sus parámetros.
        /// </summary>
        public System.Data.DataSet TraerDataset(string ProcedimientoAlmacenado, params System.Object[] Args)
        {
            System.Data.DataSet mDataSet = new System.Data.DataSet();
            this.CrearDataAdapter(ProcedimientoAlmacenado, Args).Fill(mDataSet);
            return mDataSet;
        }

        /// <summary>
        /// Obtiene un DataTable a partir de una sentencia SQL.
        /// </summary>
        public System.Data.DataTable TraerDataTable(object strSQL)
        {
            try
            {
                return TraerDataset(strSQL).Tables[0].Copy();
            }
            catch (Exception e)
            {
                // Crea datatable con el numero de error y el mensaje
                //return null;
                throw e;
            }
            
        }

        /// <summary>
        /// Obtiene un DataTable a partir de un Procedimiento Almacenado.
        /// </summary>
        public System.Data.DataTable TraerDataTable(string ProcedimientoAlmacenado)
        {
            return TraerDataset(ProcedimientoAlmacenado).Tables[0].Copy();
        }

        /// <summary>
        /// Obtiene un DataTable a partir de un Procedimiento Almacenado y sus parámetros.
        /// </summary>
        public System.Data.DataTable TraerDataTable(string ProcedimientoAlmacenado, params System.Object[] Args)
        {
            return TraerDataset(ProcedimientoAlmacenado, Args).Tables[0].Copy();
        }

        /// <summary>
        /// Obtiene un Valor a partir de una sentencia SQL.
        /// </summary>
        public System.Object TraerValor(object strSQL)
        {
            System.Data.IDbCommand Com = Comando(strSQL);
            System.Object Resp = null;

            Resp = Com.ExecuteScalar();
           
            //foreach (System.Data.IDbDataParameter Par in Com.Parameters)
            //    if (Par.Direction == System.Data.ParameterDirection.InputOutput || Par.Direction == System.Data.ParameterDirection.Output || Par.Direction == System.Data.ParameterDirection.ReturnValue)
            //        Resp = Par.Value;
            return Resp;
        }

        /// <summary>
        /// Obtiene un Valor a partir de un Procedimiento Almacenado.
        /// </summary>
        public System.Object TraerValor(string ProcedimientoAlmacenado)
        {
            System.Data.IDbCommand Com = Comando(ProcedimientoAlmacenado);  
            System.Object Resp = null;

            Resp = Com.ExecuteScalar();

            //foreach (System.Data.IDbDataParameter Par in Com.Parameters)
            //    if (Par.Direction == System.Data.ParameterDirection.InputOutput || Par.Direction == System.Data.ParameterDirection.Output || Par.Direction == System.Data.ParameterDirection.ReturnValue)
            //        Resp = Par.Value;
            return Resp;
        }

        /// <summary>
        /// Obtiene un Valor a partir de un Procedimiento Almacenado, y sus parámetros.
        /// </summary>
        public System.Object TraerValor(string ProcedimientoAlmacenado, params System.Object[] Args)
        {
            System.Data.IDbCommand Com = Comando(ProcedimientoAlmacenado);
            CargarParametros(Com, Args);

            System.Object Resp = null;

            Resp = Com.ExecuteScalar();
            //foreach (System.Data.IDbDataParameter Par in Com.Parameters)
            //    if (Par.Direction == System.Data.ParameterDirection.InputOutput || Par.Direction == System.Data.ParameterDirection.Output || Par.Direction == System.Data.ParameterDirection.ReturnValue)
            //        Resp = Par.Value;
            return Resp;
        }

        #endregion

        #region " Acciones "

        /// <summary>
        /// Interface crear conexión
        /// </summary>
        protected abstract System.Data.IDbConnection CrearConexion(string Cadena);

        /// <summary>
        /// Interface comando. Procedimieno almacenado.
        /// </summary>
        protected abstract System.Data.IDbCommand Comando(string ProcedimientoAlmacenado);

        /// <summary>
        /// Interface comando. Sentencia SQL.
        /// </summary>
        protected abstract System.Data.IDbCommand Comando(object strSQL);

        /// <summary>
        /// Interface crear dataAdapter. Procedimiento almacenado
        /// </summary>
        protected abstract System.Data.IDataAdapter CrearDataAdapter(string ProcedimientoAlmacenado, params System.Object[] Args);

        /// <summary>
        /// Interface crear dataAdapter. Sentencia SQL
        /// </summary>
        protected abstract System.Data.IDataAdapter CrearDataAdapter(object strSQL);

        /// <summary>
        /// Método cargar parámetros
        /// </summary>
        protected abstract void CargarParametros(System.Data.IDbCommand Comando, System.Object[] Args);

        // Método de extensión de la clase abstracta, debe implementarse en cada una de las clases derivadas, DatosSQLServer, DatosExcel, etc.
        // IAR mier 09DIC2015
        // link consulta
        // http://stackoverflow.com/questions/13783362/how-to-create-new-methods-in-classes-derived-from-an-abstract-class
        public abstract Object ObtieneListaTablas();

        //OPOP*********************************
        // IAR mier 21NOV2018
        // Método de extensión de la clase abstracta
        // Regresa la lista de tablas/hojas de una base de datos/Archivo excel
        // Se implementa según las necesidades
        //////public abstract Object ListaTablas();

        // IAR mier 21NOV2018
        // Método de extensión de la clase abstracta
        // Regresa la lista de columnas de una tabla de BD/Archivo excel
        // Se implementa según las necesidades
        //public abstract Object ObtieneListaColumnas(String NombreTabla);
        public abstract Object ObtieneListaColumnas(String Ruta, String NomArchivo, String NombreTabla);
        
        //OPOP*********************************

        /// <summary>
        /// Método cerrar conexión
        /// </summary>
        public Boolean CerrarConexion()
        {
            Boolean ConexionCerrada = false;

            if (mConexion == null)
            {
                // no puede cerrar la conexion ya que es nula
                ConexionCerrada = false;
            }
            else
            {
                switch (mConexion.State)
                {
                    // revisa en que estado se encuentra la conexión
                    case System.Data.ConnectionState.Broken:
                        mConexion.Close();
                        mConexion.Dispose();
                        ConexionCerrada = true;
                        break;

                    case System.Data.ConnectionState.Closed:
                        mConexion.Dispose();
                        ConexionCerrada = true;
                        break;

                    case System.Data.ConnectionState.Connecting:
                        ConexionCerrada = false;
                        break;

                    case System.Data.ConnectionState.Executing:
                        ConexionCerrada = false;
                        break;

                    case System.Data.ConnectionState.Fetching:
                        ConexionCerrada = false;
                        break;

                    case System.Data.ConnectionState.Open:
                        if (EnTransaccion)
                        {
                            ConexionCerrada = false;
                        }
                        else
                        {
                            mConexion.Close();
                            mConexion.Dispose();
                            ConexionCerrada = true;
                        }
                        break;
                }
            }

            return ConexionCerrada;
            
        }

        /// <summary>
        /// Ejecuta una sentencia de SQL en la base de datos.
        /// </summary>
        public int Ejecutar(object strSQL)
        {

            return Comando(strSQL).ExecuteNonQuery();
        }

        /// <summary>
        /// Ejecuta un Procedimiento Almacenado en la base.
        /// </summary>
        public int Ejecutar(string ProcedimientoAlmacenado)
        {
            return Comando(ProcedimientoAlmacenado).ExecuteNonQuery();
        }

        /// <summary>
        /// Ejecuta un Procedimiento Almacenado en la base, utilizando los parámetros.
        /// </summary>
        public int Ejecutar(string ProcedimientoAlmacenado, params  System.Object[] Args)
        {
            System.Data.IDbCommand Com = Comando(ProcedimientoAlmacenado);
            CargarParametros(Com, Args);
            int Resp = Com.ExecuteNonQuery();
            for (int i = 0; i < Com.Parameters.Count; i++)
            {
                System.Data.IDbDataParameter Par = (System.Data.IDbDataParameter)Com.Parameters[i];
                if (Par.Direction == System.Data.ParameterDirection.InputOutput || Par.Direction == System.Data.ParameterDirection.Output)
                    Args.SetValue(Par.Value, i - 1);
            }
            return Resp;
        }

        #endregion

        #region " Transacciones "

        protected System.Data.IDbTransaction mTransaccion;
        protected bool EnTransaccion = false;

        /// <summary>
        /// Comienza una Transacción en la base en uso.
        /// </summary>
        public void IniciarTransaccion()
        {
            mTransaccion = this.Conexion.BeginTransaction();
            EnTransaccion = true;
        }

        /// <summary>
        /// Confirma la transacción activa.
        /// </summary>
        public void TerminarTransaccion()
        {
            try
            {
                mTransaccion.Commit();
            }
            catch (System.Exception Ex)
            {
                throw Ex;
            }
            finally
            {
                mTransaccion = null;
                EnTransaccion = false;
            }
        }

        /// <summary>
        /// Cancela la transacción activa.
        /// </summary>
        public void AbortarTransaccion()
        {
            try
            {
                mTransaccion.Rollback();
            }
            catch (System.Exception Ex)
            {
                throw Ex;
            }
            finally
            {
                mTransaccion = null;
                EnTransaccion = false;
            }
        }

        #endregion

    }

}
