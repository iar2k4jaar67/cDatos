
// IAR dom 10JUL2022
// texto agreago para subir cambios en github

using System;
using System.Collections.Generic;
using System.Windows.Forms;


namespace TestData
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void menuItem4_Click(object sender, EventArgs e)
        {
            {
                // usa autenticación de sql server
                //Datos.cDatos.gDatos Serv = new Datos.cDatos.DatosSQLServer("ISI-APPSERVER", "Scorecard_SalesComp","garcileo","23456");
                //dgvResult.DataSource = Serv.TraerDataset("spGetAllMetrics").Tables[0];

                // usa autenticacion de windows
                // los valores de servidor, base de datos, usuario y password deben ser leidos a traves de settings
                Datos.cDatos.gDatos Serv = new Datos.cDatos.DatosSQLServer("GVS00993,2048", "DM_SCO_TEST", "", "");
                Serv.Autenticacion = Datos.cDatos.gDatos.TipoAutenticacion.WindowsAutentication;
                
                // uso de transacciones
                Serv.IniciarTransaccion();
                
                try
	            {	
                    //object strSQL;
                    //strSQL = "SELECT * FROM [CATALOG - Plan Frequencies]";
		            dgvResult.DataSource = Serv.TraerDataset("spCons_ListDBNames").Tables[0];
                    //dgvResult.DataSource = Serv.TraerDataset(strSQL).Tables[0];
                    Serv.TerminarTransaccion();

                    if (!Serv.CerrarConexion()) 
                    {
                        MessageBox.Show("No se puede cerrar la conexión", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

	            }
	            catch (Exception ex)
	            {
                    MessageBox.Show(ex.Message, "Validación", MessageBoxButtons.OK, MessageBoxIcon.Error);
		            Serv.AbortarTransaccion();
	            }

            }

        }

        private void menuItem2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void menuItem6_Click(object sender, EventArgs e)
        {
            // usa autenticacion de windows
            // los valores de servidor, base de datos, usuario y password deben ser leidos a traves de settings
            Datos.cDatos.gDatos Serv = new Datos.cDatos.DatosSQLServer("GVS00993,2048", "DM_SCO_TEST", "", "");
            Serv.Autenticacion = Datos.cDatos.gDatos.TipoAutenticacion.WindowsAutentication;

            // uso de transacciones
            Serv.IniciarTransaccion();

            try
            {
                dgvResult.DataSource = Serv.TraerDataset("spCons_GetExecutionList","DM_SCO_TEST").Tables[0];
                Serv.TerminarTransaccion();

                if (!Serv.CerrarConexion())
                {
                    MessageBox.Show("No se puede cerrar la conexión", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Validación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Serv.AbortarTransaccion();
            }
        }

        private void menuItem7_Click(object sender, EventArgs e)
        {
            // usa autenticacion de windows
            // los valores de servidor, base de datos, usuario y password deben ser leidos a traves de settings
            Datos.cDatos.gDatos Serv = new Datos.cDatos.DatosSQLServer("GVS00993,2048", "DM_SCO_TEST", "", "");
            Serv.Autenticacion = Datos.cDatos.gDatos.TipoAutenticacion.WindowsAutentication;
  
            // uso de transacciones
            //Serv.IniciarTransaccion();

            try
            {
                object objSQL = "";
                string strSQL = "";
                objSQL ="select count(*) from [CATALOG - Plan Frequencies]";
                strSQL = "spCons_GetAvailable_FY";
                System.Windows.Forms.MessageBox.Show(Serv.TraerValor(strSQL).ToString(), "Valor", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //Serv.TerminarTransaccion();
            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.Message, "Validación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //Serv.AbortarTransaccion();
            }

//            -- un solo valor query
//select plan_frequency from [CATALOG - Plan Frequencies] where plan_frequency = 'annual'

//spCons_GetExecutionList 'DM_SCO_TEST'

//-- un solo valor sp sin parametros
//spCons_GetAvailable_FY
        }

        private void menuItem3_Click(object sender, EventArgs e)
        {

            //Datos.cDatos.gDatos Serv = new Datos.cDatos.DatosOracle("QPODSP.americas.hpqcorp.net", "", "s_tables", "tables");
            //Serv.Autenticacion = Datos.cDatos.gDatos.TipoAutenticacion.WindowsAutentication;

            // uso de transacciones
            //Serv.IniciarTransaccion();

            object objSQL = "";
            objSQL = "select * from US_OMEGA_TRANS_DTL where rownum < 10000";

            try
            {
                //Serv.TimeOut = 90;

                //dgvResult.DataSource = Serv.TraerDataset(objSQL).Tables[0];
                //Serv.TerminarTransaccion();

                //if (!Serv.CerrarConexion())
                //{
                //    MessageBox.Show("No se puede cerrar la conexión", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Validación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //Serv.AbortarTransaccion();
            }
        }

        private void menuItem5_Click(object sender, EventArgs e)
        {

            object objSQL = "";
            objSQL = "SELECT * FROM [hoja1$]";

            Datos.cDatos.gDatos Serv = new Datos.cDatos.DatosExcel("C:\\Boxito\\", "PruebaBD.xlsx", "", "");
            ////Serv.Autenticacion = Datos.cDatos.gDatos.TipoAutenticacion.WindowsAutentication;

            //////dgvResult.DataSource = Serv.TraerDataset(objSQL).Tables[0];
            System.Data.DataTable odt = new System.Data.DataTable();
            odt = Serv.TraerDataTable(objSQL);
            dgvResult.DataSource = odt;


        }

        private void menuItem9_Click(object sender, EventArgs e)
        {

            Datos.cDatos.gDatos Serv = new Datos.cDatos.DatosExcel("C:\\Boxito\\", "PruebaBD.xlsx", "", "");
            
            List<String> lstHojas = new List<String>();
            Object obj = new Object();

            obj = Serv.ObtieneListaTablas();
            lstHojas = (List<String>)obj;

        }

        private void menuItem10_Click(object sender, EventArgs e)
        {
            
            Datos.cDatos.gDatos Serv = new Datos.cDatos.DatosExcel("C:\\Boxito\\", "PruebaBD.xlsx", "", "");

            List<string> lstHojas = new List<string>();

            List<string> lstCols = new List<string>();

            lstHojas = (List<string>)Serv.ObtieneListaTablas();

            foreach (String str in lstHojas)
            {
                if (str == "bd3$")
                {
                    
                    lstCols = (List<string>)Serv.ObtieneListaColumnas("C:\\Boxito\\", "PruebaBD.xlsx", str.ToString());
                }
            }

            
            
        }


    }
}