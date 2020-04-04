using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.IO;
using System.Timers;
using SAPbobsCOM;
using Interfaz_MADEGOM.Functions;
using System.Globalization;
using System.Net;
using System.Xml;
//using System.Core;
using Microsoft.CSharp;
using WinSCP;


namespace Interfaz_MADEGOM
{
    public partial class Service1 : ServiceBase
    {
        //public Timer Tiempo; luego decomento
        public String s;
        public SAPbobsCOM.CompanyClass oCompany;
        public SAPbobsCOM.Recordset oRecordSet;
        public CultureInfo _nf = new System.Globalization.CultureInfo("en-US");
        public String[] arr;
        public String PathArchivos;
        public String PathOK;
        public String PathNoOk;
        public String PathResp;
        public System.Data.DataTable dtlote;
        public System.Data.DataTable dtsls;

        /// <summary>
        /// Constructor y asignancion del intervalo de bajada
        /// </summary>
        public Service1()
        {

            InitializeComponent();
           /* Tiempo = new Timer();
            Tiempo.Interval = 30000;
            Tiempo.Elapsed += new ElapsedEventHandler(tiempo_elapsed);
            */

        }

        protected override void OnStart(string[] args)
        {
            Tiempo.Enabled = true;
        }

        protected override void OnStop()
        {
            
            Tiempo.Stop();
            Tiempo.Enabled = false;
             
        }


        //metodo principal (inicio)
        private void Tiempo_Elapsed(object sender, ElapsedEventArgs e)
        {
            TFunctions Func;
            System.IO.StreamReader sr;
            Int32 i;
            String sLine = "";
            String Tipo = "";
            String[] aline;
            String[] files;
            System.Data.DataTable dt;
            System.Data.DataTable dtSorted;
            System.Data.DataColumn column;
            String CardCode = "";
            String WhsCode = "";
            DateTime DocDate = DateTime.Now;
            String Docs;
            DataRow dtrow;
            DataRow dtrowsls;
            DataRow[] foundRows;
            Boolean bPasoOk;
            Int32 iCont;
            Boolean bSeProcesa = false;
            Boolean bCerrarArchivo = false;
            String sDocEntry = "";
            Boolean bAgregarLinea = true;
            String Code = "";
            Double dB99;
            Double dB07;
            Double d02;

            Func = new TFunctions();
            Tiempo.Stop();
            Tiempo.Enabled = false;
            oCompany = new SAPbobsCOM.CompanyClass();
            try
            {
                Func.AddLog("Inicio");
                Download();  //*

                if (ConectarBaseSAP())
                {
                    Func.AddLog("Conectado a SAP");


                    oRecordSet = (SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
                    dtlote = new System.Data.DataTable();
                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.String");
                    column.ColumnName = "ItemCode";
                    dtlote.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = Type.GetType("System.String");
                    column.ColumnName = "Lote";
                    dtlote.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = Type.GetType("System.String");
                    column.ColumnName = "Status";
                    dtlote.Columns.Add(column);

                    dtsls = new System.Data.DataTable();
                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.String");
                    column.ColumnName = "DocEntry";
                    dtsls.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.String");
                    column.ColumnName = "CardCode";
                    dtsls.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.String");
                    column.ColumnName = "Carton";
                    dtsls.Columns.Add(column);

                    //dtsls = new System.Data.DataTable();
                    //column = new DataColumn();
                    //column.DataType = System.Type.GetType("System.String");
                    //column.ColumnName = "Carton";
                    //dtsls.Columns.Add(column);

                    

                    arr = new System.String[10];
                    arr = Func.IniciaArchivos();
                    dt = new System.Data.DataTable();
                    dtSorted = new System.Data.DataTable();

                    if (arr != null)
                    {
                        i = 0;
                        if (i < arr.Length)
                        {

                            
                            PathArchivos = Func.PathArchivos();
                            
                            if (Directory.Exists(PathArchivos))
                            {
                                PathOK = PathArchivos + "\\ProcesadoOK";
                                PathNoOk = PathArchivos + "\\ProcesadoError";
                                PathResp = Func.PathArchivosResp();

                                if (!Directory.Exists(PathOK))
                                    Directory.CreateDirectory(PathOK);
                                if (!Directory.Exists(PathNoOk))
                                    Directory.CreateDirectory(PathNoOk);
                                if (!Directory.Exists(PathResp))
                                    Directory.CreateDirectory(PathResp);

                                files = System.IO.Directory.GetFiles(PathArchivos);
                                foreach (String sv in files)
                                {

                                    //Func.AddLog("Paso x3");
                                    Tipo = "";
                                    CardCode = "";
                                    DocDate = DateTime.Now;
                                    WhsCode = "";
                                    Docs = sv.Substring(PathArchivos.Length, 3);            //productivo 
                                //    Docs = sv.Substring(PathArchivos.Length+1, 3);             //** se le asingo +1 ya que estaba fallando probar en
                                    bPasoOk = false;
                                    i = 0;
                                    bSeProcesa = false;
                                    Func.AddLog("Docs: " +Docs);
                                    while (i < arr.Length)
                                    {
                                        if (Docs == arr[i])
                                        { 
                                            bSeProcesa = true;
                                            i = arr.Length;
                                        }
                                        i++;
                                    }

                                    if (bSeProcesa)
                                    {
                                        Func.AddLog(sv);
                                        dtSorted.Clear();
                                        dtSorted.Columns.Clear();
                                        dt.Clear();
                                        dt.Columns.Clear();

                                        dtlote.Rows.Clear();
                                        dtsls.Rows.Clear();
                                        sr = new System.IO.StreamReader(sv);
                                        try
                                        {
                                            bCerrarArchivo = true;
                                            iCont = 0;
                                            sLine = "";
                                            //texto := sr.ReadToEnd();
                                            while (sLine != null)
                                            {
                                                //Func.AddLog("Paso x4");
                                                sLine = sr.ReadLine();
                                                if (sLine != null)
                                                {
                                                    if (sLine != "")
                                                    {
                                                        aline = sLine.Split('|');
                                                        if (Docs == "INS") //para resumen de inventario
                                                        {
                                                            bAgregarLinea = true;
                                                            if (iCont == 0)
                                                            {
                                                                column = new DataColumn();
                                                                column.DataType = System.Type.GetType("System.String");
                                                                column.ColumnName = "ItemCode";
                                                                dt.Columns.Add(column);

                                                                column = new DataColumn();
                                                                column.DataType = System.Type.GetType("System.String");
                                                                column.ColumnName = "WhsCode";
                                                                dt.Columns.Add(column);

                                                                column = new DataColumn();
                                                                column.DataType = System.Type.GetType("System.String");
                                                                column.ColumnName = "Lote";
                                                                dt.Columns.Add(column);

                                                                column = new DataColumn();
                                                                column.DataType = System.Type.GetType("System.String");
                                                                column.ColumnName = "UM";
                                                                dt.Columns.Add(column);

                                                                column = new DataColumn();
                                                                column.DataType = Type.GetType("System.Double");
                                                                column.ColumnName = "Quantity";
                                                                dt.Columns.Add(column);
                                                            }

                                                            dB99 = Convert.ToDouble(aline[38].Replace(",", "."), _nf);
                                                            dB07 = Convert.ToDouble(aline[40].Replace(",", "."), _nf);
                                                            d02 = Convert.ToDouble(aline[35].Replace(",", "."), _nf) - dB99 - dB07;

                                                            dtrow = dt.NewRow();
                                                            dtrow["ItemCode"] = aline[4];
                                                            dtrow["UM"] = "";
                                                            dtrow["WhsCode"] = "99";
                                                            dtrow["Lote"] = aline[15];
                                                            dtrow["Quantity"] = dB99;
                                                            dt.Rows.Add(dtrow);

                                                            dtrow = dt.NewRow();
                                                            dtrow["ItemCode"] = aline[4];
                                                            dtrow["UM"] = "";
                                                            dtrow["WhsCode"] = "07";
                                                            dtrow["Lote"] = aline[15];
                                                            dtrow["Quantity"] = dB07;
                                                            dt.Rows.Add(dtrow);

                                                            dtrow = dt.NewRow();
                                                            dtrow["ItemCode"] = aline[4];
                                                            dtrow["UM"] = "";
                                                            dtrow["WhsCode"] = "02";
                                                            dtrow["Lote"] = aline[15];
                                                            dtrow["Quantity"] = d02;
                                                            dt.Rows.Add(dtrow);

                                                        }
                                                        else
                                                        {
                                                            if (Docs == "IHT") //Historial
                                                            {
                                                                bAgregarLinea = true;
                                                                if (iCont == 0)
                                                                {
                                                                    column = new DataColumn();
                                                                    column.DataType = System.Type.GetType("System.String");
                                                                    column.ColumnName = "Code";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = System.Type.GetType("System.String");
                                                                    column.ColumnName = "CodBloqueo";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.String");
                                                                    column.ColumnName = "Razon";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.Double");
                                                                    column.ColumnName = "Cantidad";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.String");
                                                                    column.ColumnName = "Ref1";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.String");
                                                                    column.ColumnName = "Ref2";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.String");
                                                                    column.ColumnName = "Ref4";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.String");
                                                                    column.ColumnName = "Orden";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.String");
                                                                    column.ColumnName = "ItemCode";
                                                                    dt.Columns.Add(column);


                                                                }

                                                                Code = aline[4];
                                                                if ((Code == "17") || (Code == "23") || (Code == "25") || (Code == "27") || (Code == "39") || (Code == "40") || (Code == "53") || (Code == "56")) //
                                                                {
                                                                    dtrow = dt.NewRow();
                                                                    dtrow["Code"] = Code;
                                                                    Func.AddLog("Code - " + Code);

                                                                    if ((aline[6] != "") && (Code == "40") && (aline[6].Length > 2))
                                                                        dtrow["CodBloqueo"] = aline[6].Substring(0,3);
                                                                    else
                                                                        dtrow["CodBloqueo"] = aline[6];
                                                                    Func.AddLog("CodBloqueo - " + aline[6]);
                                                                    
                                                                    dtrow["Razon"] = aline[5];
                                                                    Func.AddLog("Razon - " + aline[5]);
                                                                    if ((Code == "23") || (Code == "25") || (Code == "40")) //|| (Code == "39")
                                                                    {
                                                                        dtrow["Cantidad"] = Convert.ToDouble(aline[26].Replace(",", "."), _nf);
                                                                        Func.AddLog("Cantidad - " + aline[26]);//cantidad origen
                                                                    }
                                                                    else if (Code == "56")
                                                                    {
                                                                        dtrow["Cantidad"] = 0;
                                                                        Func.AddLog("Cantidad - " + aline[27]);
                                                                    }
                                                                    else
                                                                    {
                                                                        dtrow["Cantidad"] = Convert.ToDouble(aline[27].Replace(",", "."), _nf);
                                                                        Func.AddLog("Cantidad - " + aline[27]);
                                                                    }
                                                                    dtrow["Ref1"] = aline[33];
                                                                    Func.AddLog("Ref1 - " + aline[33]);
                                                                    dtrow["Ref2"] = aline[35];
                                                                    Func.AddLog("Ref2 - " + aline[35]);
                                                                    dtrow["Ref4"] = aline[39];
                                                                    Func.AddLog("Ref4 - " + aline[39]);
                                                                    dtrow["Orden"] = aline[23];
                                                                    Func.AddLog("Orden - " + aline[23]);
                                                                    dtrow["ItemCode"] = aline[9];
                                                                    Func.AddLog("ItemCode - " + aline[9]);
                                                                    dt.Rows.Add(dtrow);
                                                                }
                                                            }
                                                            else if (aline[0].Substring(0, 4) == "[H1]")
                                                            {
                                                                if (Docs == "SVS")
                                                                {
                                                                    Tipo = aline[5];
                                                                    CardCode = aline[9];
                                                                    DocDate = DateTime.ParseExact(aline[20].Substring(0, 8), "yyyyMMdd", CultureInfo.InvariantCulture);
                                                                    WhsCode = aline[1];

                                                                    column = new DataColumn();
                                                                    column.DataType = System.Type.GetType("System.String");
                                                                    column.ColumnName = "ItemCode";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.Double");
                                                                    column.ColumnName = "Quantity";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.String");
                                                                    column.ColumnName = "Lote";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.DateTime");
                                                                    column.ColumnName = "ExpDate";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.Int32");
                                                                    column.ColumnName = "ObjType";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.Int32");
                                                                    column.ColumnName = "DocEntry";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.Int32");
                                                                    column.ColumnName = "LineNum";
                                                                    dt.Columns.Add(column);

                                                                }

                                                                if (Docs == "SLS")
                                                                {
                                                                    Tipo = aline[2]; //debe ser CREATE para funcionar

                                                                    column = new DataColumn();
                                                                    column.DataType = System.Type.GetType("System.String");
                                                                    column.ColumnName = "Address";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.String");
                                                                    column.ColumnName = "County";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.String");
                                                                    column.ColumnName = "City";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.String");
                                                                    column.ColumnName = "State";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.Int32");
                                                                    column.ColumnName = "ObjType";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.Int32");
                                                                    column.ColumnName = "DocEntry";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.Int32");
                                                                    column.ColumnName = "LineNum";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.String");
                                                                    column.ColumnName = "ItemCode";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.Double");
                                                                    column.ColumnName = "Quantity";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.String");
                                                                    column.ColumnName = "CardCode";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.String");
                                                                    column.ColumnName = "Lote";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.DateTime");
                                                                    column.ColumnName = "ExpDate";
                                                                    dt.Columns.Add(column);

                                                                    column = new DataColumn();
                                                                    column.DataType = Type.GetType("System.String");
                                                                    column.ColumnName = "Carton";
                                                                    dt.Columns.Add(column);

                                                                }
                                                            }
                                                            else
                                                            {
                                                                //una vez creada la estructura del datatable hay que llenarlo
                                                                dtrow = dt.NewRow();
                                                                if (Docs == "SVS")
                                                                {
                                                                    if (Tipo != "")
                                                                    {
                                                                        if (aline[1] != "")
                                                                        {
                                                                            if (aline[1].Substring(0, 3) == "LPN")
                                                                            {
                                                                                bAgregarLinea = true;
                                                                                dtrow["ItemCode"] = aline[5];
                                                                                if (aline[18].Replace(",", ".") == "")
                                                                                    throw new Exception("No se encuentra Cantidad");
                                                                                dtrow["Quantity"] = Convert.ToDouble(aline[18].Replace(",", "."), _nf);
                                                                                dtrow["Lote"] = aline[25];
                                                                                if (aline[24] != "")
                                                                                    dtrow["ExpDate"] = DateTime.ParseExact(aline[24].Substring(0, 8), "yyyyMMdd", CultureInfo.InvariantCulture);
                                                                                if (aline[27] == "")
                                                                                    throw new Exception("No se encuentra ObjType");
                                                                                dtrow["ObjType"] = Convert.ToInt32(aline[27]);
                                                                                if (aline[28] == "")
                                                                                    throw new Exception("No se encuentra DocEntry");
                                                                                dtrow["DocEntry"] = Convert.ToInt32(aline[28]);
                                                                                if (aline[29] == "")
                                                                                    throw new Exception("No se encuentra LineNum");
                                                                                dtrow["LineNum"] = Convert.ToInt32(aline[29]);
                                                                            }
                                                                            else
                                                                                bAgregarLinea = false;
                                                                        }
                                                                        else
                                                                            bAgregarLinea = false;
                                                                    }
                                                                }
                                                                else if (Docs == "SLS")
                                                                {
                                                                    bAgregarLinea = true;
                                                                    //CardCode = aline[79];
                                                                    dtrow["Address"] = aline[48]; // +" " + aline[28];
                                                                    dtrow["County"] = aline[29];
                                                                    dtrow["City"] = aline[30];
                                                                    dtrow["State"] = aline[31];
                                                                    if ((aline[51] == "0") || (aline[51] == "")) //para poder hacer las transferencias a BOD52 y BOD54
                                                                    {
                                                                        dtrow["ObjType"] = 0;
                                                                        dtrow["DocEntry"] = 0;
                                                                        dtrow["LineNum"] = 0;
                                                                        dtrow["CardCode"] = aline[20];
                                                                    }
                                                                    else
                                                                    {
                                                                        dtrow["ObjType"] = Convert.ToInt32(aline[50]);
                                                                        dtrow["DocEntry"] = Convert.ToInt32(aline[51]);
                                                                        dtrow["LineNum"] = Convert.ToInt32(aline[52]);
                                                                        dtrow["CardCode"] = aline[47];
                                                                    }

                                                                    dtrow["ItemCode"] = aline[56];
                                                                    if (aline[72].Replace(",", ".") == "")
                                                                        throw new Exception("No se encuentra Cantidad");
                                                                    dtrow["Quantity"] = Convert.ToDouble(aline[72].Replace(",", "."), _nf);
                                                                    dtrow["Lote"] = aline[75];
                                                                    if (aline[76] != "")
                                                                        dtrow["ExpDate"] = DateTime.ParseExact(aline[76].Substring(0, 8), "yyyyMMdd", CultureInfo.InvariantCulture);

                                                                    dtrow["Carton"] = aline[55];
                                                                    //if (sDocEntry != aline[51])
                                                                    //{
                                                                    //    dtrowsls = dtsls.NewRow();
                                                                    //    dtrowsls["DocEntry"] = aline[51];
                                                                    //    dtrowsls["CardCode"] = aline[47];
                                                                    //    dtsls.Rows.Add(dtrowsls);
                                                                    //}
                                                                    sDocEntry = aline[51];
                                                                }

                                                                if (bAgregarLinea)
                                                                    dt.Rows.Add(dtrow);
                                                            }
                                                        }
                                                    }
                                                }
                                                iCont++;
                                            }
                                            sr.Close();//se cierra el archivo para moverlo
                                            bCerrarArchivo = false;
                                            dtSorted.Clear();

                                            if ((Docs != "INS") && (Docs != "IHT"))
                                            {
                                                //ordenar datatable por documento base y linenum
                                                foundRows = dt.Select("", "ObjType ASC, DocEntry ASC, LineNum ASC");
                                                if (foundRows.Length == 0)
                                                    throw new Exception("No se encuentra lineas a procesar (considerar que deben tener LPN)");
                                                dtSorted = foundRows.CopyToDataTable();
                                            }

                                            if ((Docs == "SVS") && (Tipo == "FAB"))
                                                bPasoOk = SVS_Transferencia(Tipo, CardCode, DocDate, WhsCode, ref dtSorted, sv);
                                            else if ((Docs == "SVS") && (Tipo != ""))
                                                bPasoOk = SVS_Documents(Tipo, CardCode, DocDate, WhsCode, ref dtSorted, sv);
                                            else if ((Docs == "SLS") && (Tipo == "CREATE"))
                                            {
                                                System.Int32 DocE = -1;
                                                System.String Cliente = "";
                                                //foreach (System.Data.DataRow ors2 in dtSorted.Rows)
                                                //{
                                                //    Func.AddLog("dtSorted " + ors2["DocEntry"].ToString() + "-" + ors2["CardCode"].ToString());
                                                //}
                                                foreach (System.Data.DataRow ors in dtSorted.Rows)
                                                {
                                                    if ((DocE != (System.Int32)ors["DocEntry"]) && ((System.Int32)ors["DocEntry"] != 0))
                                                    {
                                                        if ((System.Int32)ors["DocEntry"] != 0)
                                                        {
                                                            dtrowsls = dtsls.NewRow();
                                                            dtrowsls["DocEntry"] = (System.Int32)ors["DocEntry"];
                                                            dtrowsls["CardCode"] = (System.String)ors["CardCode"];
                                                            dtrowsls["Carton"] = (System.String)ors["Carton"];
                                                            dtsls.Rows.Add(dtrowsls);
                                                        }
                                                    }
                                                    else if (((System.Int32)ors["DocEntry"] == 0) && (Cliente != (System.String)ors["CardCode"]))
                                                    {
                                                        dtrowsls = dtsls.NewRow();
                                                        dtrowsls["DocEntry"] = (System.Int32)ors["DocEntry"];
                                                        dtrowsls["CardCode"] = (System.String)ors["CardCode"];
                                                        dtrowsls["Carton"] = (System.String)ors["Carton"];
                                                        dtsls.Rows.Add(dtrowsls);
                                                    }
                                                    DocE = (System.Int32)ors["DocEntry"];
                                                    if ((System.Int32)ors["DocEntry"] == 0)
                                                        Cliente = (System.String)ors["CardCode"];
                                                }

                                                //foreach (System.Data.DataRow ors1 in dtsls.Rows)
                                                //{
                                                //    Func.AddLog("dtsls " + ors1["DocEntry"].ToString() + "-" + ors1["CardCode"].ToString());
                                                //}
                                                bPasoOk = SLS_Documents(Tipo, ref dtsls, ref dtSorted, sv);
                                            }
                                            else if (Docs == "INS") //para resumen de inventario
                                                bPasoOk = INS_ResumenInventario(ref dt, sv);
                                            else if (Docs == "IHT") //para historial de inventario
                                                bPasoOk = IHT_Documents(ref dt, sv);


                                            if (bPasoOk)
                                            {
                                                //para copiar archivo en direccion de respaldo
                                                s = PathResp + "\\" + sv.Substring(PathArchivos.Length, sv.Length - PathArchivos.Length);
                                                if (File.Exists(s))
                                                    File.Delete(s);
                                                File.Copy(sv, s);

                                                //para mover a carpeta de procesado correctamente
                                                s = PathOK + "\\" + sv.Substring(PathArchivos.Length, sv.Length - PathArchivos.Length);
                                                if (File.Exists(s))
                                                    File.Delete(s);
                                                File.Move(sv, s);
                                            }
                                            else
                                            {
                                                //para mover a carpeta de procesado con error
                                                s = PathNoOk + "\\" + sv.Substring(PathArchivos.Length, sv.Length - PathArchivos.Length);
                                                if (File.Exists(s))
                                                    File.Delete(s);
                                                File.Move(sv, s);
                                            }

                                        }
                                        catch (Exception q)
                                        {
                                            Func.AddLog("Procesar archivo " + sv + ": " + q.Message + " ** Trace: " + q.StackTrace);
                                            if (bCerrarArchivo)
                                                sr.Close();
                                        }
                                        finally
                                        {
                                            ;
                                        }
                                    }


                                }//fin foreach
                            }
                            else
                                Func.AddLog("No se encuentra directorio con los archivos, " + PathArchivos);
                            //Func.AddLog("Paso 8");
                        }
                    }
                    else
                        Func.AddLog("No se encuentran el incicio de archivos a considerar en Config.xml");

                    if (oCompany != null)
                        oCompany.Disconnect();
                    oCompany = null;
                    oRecordSet = null;
                }
                else
                    Func.AddLog("No se ha podido conectar a la Base SAP, revisar datos de conexion");//no se ha podido conectar
            }
            catch (Exception w)
            {
                Func.AddLog("timer1_elapsed: " + w.Message + " ** Trace: " + w.StackTrace);
            }
            finally
            {
                Tiempo.Enabled = true;
                Tiempo.Start();
                Func.AddLog("Fin");
            }
        }

        //crear objetivos
        public Boolean SVS_Transferencia(String Tipo, String CardCode, DateTime DocDate, String WhsCode, ref System.Data.DataTable dtSorted, String sv)
        {
            Boolean _return = false;
            TFunctions Func = new TFunctions();
            SAPbobsCOM.StockTransfer oTransfer;
            Int32 ilinea;
            Int32 ilotelinea;
            String ItemCode;
            Double dquantity;
            Int32 lRetCode;
            Int32 lErrcode;
            String sErrmsg;
            DataRow dtrow;
            Boolean ManejaLote;

            try
            {
                oTransfer = (SAPbobsCOM.StockTransfer)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer));
                oTransfer.CardCode = CardCode;
                oTransfer.DocDate = DocDate;
                oTransfer.FromWarehouse = "55";

                ItemCode = "";
                dquantity = 0;
                ilotelinea = 0;
                ilinea = 0;
                foreach (System.Data.DataRow orows in dtSorted.Rows)
                {
                    if (ItemCode != ((System.String)orows["ItemCode"]).Trim())
                    {
                        oTransfer.Lines.Quantity = dquantity;

                        if (ilinea > 0)
                        {
                            oTransfer.Lines.Add();
                            ilotelinea = 0;
                            dquantity = 0;
                        }

                        oTransfer.Lines.ItemCode = ((System.String)orows["ItemCode"]).Trim();
                        oTransfer.Lines.WarehouseCode = "02";
                        oTransfer.Lines.Quantity = dquantity;
                    }

                    s = "SELECT ManBtchNum FROM OITM WHERE ItemCode = '{0}'";
                    s = String.Format(s, ((System.String)orows["ItemCode"]).Trim());
                    oRecordSet.DoQuery(s);
                    if (oRecordSet.RecordCount == 0)
                        ManejaLote = false;
                    else
                        if (((System.String)oRecordSet.Fields.Item("ManBtchNum").Value).Trim() == "Y")
                            ManejaLote = true;
                        else
                            ManejaLote = false;


                    if ((((System.String)orows["Lote"]).Trim() != "") && (ManejaLote))
                    {
                        if (ilotelinea > 0)
                            oTransfer.Lines.BatchNumbers.Add();
                        if (orows["ExpDate"].ToString() != "")
                            oTransfer.Lines.BatchNumbers.ExpiryDate = ((System.DateTime)orows["ExpDate"]);
                        oTransfer.Lines.BatchNumbers.Quantity = ((System.Double)orows["Quantity"]);
                        oTransfer.Lines.BatchNumbers.BatchNumber = ((System.String)orows["Lote"]);
                        dquantity = dquantity + ((System.Double)orows["Quantity"]);
                        ilotelinea++;
                        ItemCode = ((System.String)orows["ItemCode"]).Trim();

                        //guardo lote en dt para actualizar status
                        s = @"select Status from OBTN where ItemCode = '{0}' and DistNumber = '{1}' and Status = '1'";
                        s = String.Format(s, ((System.String)orows["ItemCode"]).Trim(), ((System.String)orows["Lote"]));
                        oRecordSet.DoQuery(s);
                        if (oRecordSet.RecordCount > 0)
                        {
                            dtrow = dtlote.NewRow();
                            dtrow["ItemCode"] = ((System.String)orows["ItemCode"]).Trim();
                            dtrow["Lote"] = ((System.String)orows["Lote"]);
                            dtrow["Status"] = ((System.String)oRecordSet.Fields.Item("Status").Value).Trim();
                            dtlote.Rows.Add(dtrow);
                        }
                    }
                    else
                        dquantity = ((System.Double)orows["Quantity"]);

                    ilinea++;
                }

                oTransfer.Lines.Quantity = dquantity;
                //Cambiar Status a lotes como no accesible
                if (dtlote.Rows.Count > 0)
                    Cambiar_Status_Lote(ref dtlote, false);
                lRetCode = oTransfer.Add();  //tabla normal 
                if (lRetCode != 0)
                {
                    oCompany.GetLastError(out lErrcode, out sErrmsg);
                    Func.AddLog("Archivo " + sv + " con problemas al crear en SAP, " + sErrmsg);
                    Func.AddLogSap(oCompany, "Archivo " + sv + " con problemas al crear en SAP, " + sErrmsg, sv, dquantity, "", "", "", "", "", "", ItemCode, "", "", "");
                    _return = false;
                    s = "C:\\transfer Ma.xml";
                    oTransfer.SaveXML(s);

                    //s = PathNoOk + "\\" + sv.Substring(PathArchivos.Length,sv.Length - PathArchivos.Length);
                    //if (File.Exists(s))
                    //    File.Delete(s);
                    //File.Move(sv, s);
                }
                else
                {
                    Func.AddLog("Archivo " + sv + " creado satisfactoriamente en SAP");
                    Func.AddLogSap(oCompany, "Archivo " + sv + " creado satisfactoriamente en SAP", sv, dquantity, "", "", "", "", "", "", ItemCode, "", "", "");
                    _return = true;
                }
                if (dtlote.Rows.Count > 0)
                    Cambiar_Status_Lote(ref dtlote, true);

                return _return;
            }
            catch (Exception we)
            {
                Func.AddLog("Error SVS_Transferencia - " + we.Message + ", StackTrace " + we.StackTrace);
                return false;
            }
        }

        //crea objetos factura compra, entrega, devolucion y nota de credito
        public Boolean SVS_Documents(String Tipo, String CardCode, DateTime DocDate, String WhsCode, ref System.Data.DataTable dtSorted, String sv)
        {
            Boolean _return = false;
            TFunctions Func = new TFunctions();
            SAPbobsCOM.Documents oDocuments = null;
            SAPbobsCOM.Documents oDocuments2;
            Int32 ilinea;
            Int32 ilotelinea;
            String ItemCode;
            Double dquantity;
            Double CantxUni = 0;
            Boolean bDividir = false;
            Int32 lRetCode;
            Int32 lErrcode;
            String sErrmsg;
            String sDocEntry;
            DataRow dtrow;
            Int32 DocEntryOC = 0;
            Boolean ManejaLote;
            try
            {
                if (Tipo == "NAC")
                    oDocuments = ((SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes));
                else if (Tipo == "IMP")
                {
                    oDocuments = ((SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices));
                }
                else if (Tipo == "DGD")
                    oDocuments = ((SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oReturns));
                else if (Tipo == "DFA")
                    oDocuments = ((SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes));

                if (oDocuments != null)
                {
                    oDocuments.CardCode = CardCode;
                    oDocuments.DocDate = DocDate;
                    if (Tipo == "IMP")
                        oDocuments.ReserveInvoice = SAPbobsCOM.BoYesNoEnum.tYES;

                    ItemCode = "";
                    dquantity = 0;
                    ilotelinea = 0;
                    ilinea = 0;
                    foreach (System.Data.DataRow orows in dtSorted.Rows)
                    {
                        if (ItemCode != ((System.String)orows["ItemCode"]).Trim())
                        {
                            if ((Tipo == "DGD") || (Tipo == "DFA"))//para devolucion y nota de credito
                            {
                                if (bDividir == true)
                                {
                                    if (CantxUni != 0)
                                        oDocuments.Lines.Quantity = dquantity / CantxUni;
                                    else
                                        oDocuments.Lines.Quantity = dquantity;
                                }
                                else
                                    oDocuments.Lines.Quantity = dquantity;
                            }
                            else
                                oDocuments.Lines.Quantity = dquantity;

                            if (ilinea > 0)
                            {
                                oDocuments.Lines.Add();
                                ilotelinea = 0;
                                dquantity = 0;
                            }

                            oDocuments.Lines.ItemCode = ((System.String)orows["ItemCode"]).Trim();
                            oDocuments.Lines.WarehouseCode = WhsCode;
                            oDocuments.Lines.Quantity = dquantity;
                            oDocuments.Lines.BaseType = ((System.Int32)orows["ObjType"]);
                            oDocuments.Lines.BaseEntry = ((System.Int32)orows["DocEntry"]);
                            oDocuments.Lines.BaseLine = ((System.Int32)orows["LineNum"]);

                            if ((Tipo == "DGD") || (Tipo == "DFA"))
                            {
                                s = "select T1.UseBaseUn, O0.NumInSale from {2} T1 join OITM O0 on O0.ItemCode = T1.ItemCode where T1.DocEntry = {0} and T1.LineNum = {1}";
                                s = String.Format(s, ((System.Int32)orows["DocEntry"]), ((System.Int32)orows["LineNum"]), Tipo == "DGD" ? "DLN1" : "INV1");
                                oRecordSet.DoQuery(s);
                                if (((System.String)oRecordSet.Fields.Item("UseBaseUn").Value).Trim() == "N")
                                {
                                    CantxUni = (System.Double)(oRecordSet.Fields.Item("NumInSale").Value);
                                    bDividir = true;
                                    if (CantxUni != 0)
                                        oDocuments.Lines.Quantity = dquantity / CantxUni;
                                    else
                                        oDocuments.Lines.Quantity = dquantity;
                                }
                                else
                                    oDocuments.Lines.Quantity = dquantity;
                            }
                            else
                                oDocuments.Lines.Quantity = dquantity;

                            DocEntryOC = oDocuments.Lines.BaseEntry;
                        }

                        s = "SELECT ManBtchNum FROM OITM WHERE ItemCode = '{0}'";
                        s = String.Format(s, ((System.String)orows["ItemCode"]).Trim());
                        oRecordSet.DoQuery(s);
                        if (oRecordSet.RecordCount == 0)
                            ManejaLote = false;
                        else
                            if (((System.String)oRecordSet.Fields.Item("ManBtchNum").Value).Trim() == "Y")
                                ManejaLote = true;
                            else
                                ManejaLote = false;

                        if (((((System.String)orows["Lote"]).Trim() != "") && (Tipo != "IMP")) && (ManejaLote))
                        {
                            if (ilotelinea > 0)
                                oDocuments.Lines.BatchNumbers.Add();   //CreateMethodEnum objetos aca?
                            if (orows["ExpDate"].ToString() != "")
                                oDocuments.Lines.BatchNumbers.ExpiryDate = ((System.DateTime)orows["ExpDate"]);
                            oDocuments.Lines.BatchNumbers.Quantity = ((System.Double)orows["Quantity"]);
                            oDocuments.Lines.BatchNumbers.BatchNumber = ((System.String)orows["Lote"]).Trim();
                            dquantity = dquantity + ((System.Double)orows["Quantity"]);
                            ilotelinea++;
                            //guardo lote en dt para actualizar status
                            s = @"select Status from OBTN where ItemCode = '{0}' and DistNumber = '{1}' and Status = '1'";
                            s = String.Format(s, ((System.String)orows["ItemCode"]).Trim(), ((System.String)orows["Lote"]));
                            oRecordSet.DoQuery(s);
                            if (oRecordSet.RecordCount > 0)
                            {
                                dtrow = dtlote.NewRow();
                                dtrow["ItemCode"] = ((System.String)orows["ItemCode"]).Trim();
                                dtrow["Lote"] = ((System.String)orows["Lote"]);
                                dtrow["Status"] = ((System.String)oRecordSet.Fields.Item("Status").Value).Trim();
                                dtlote.Rows.Add(dtrow);
                            }
                        }
                        if ((((System.String)orows["Lote"]).Trim() != "") && (Tipo == "IMP"))
                            dquantity = dquantity + ((System.Double)orows["Quantity"]);
                        else if ((((System.String)orows["Lote"]).Trim() == "") || (!ManejaLote))
                            dquantity = dquantity + ((System.Double)orows["Quantity"]);
                        ItemCode = ((System.String)orows["ItemCode"]).Trim();

                        ilinea++;
                    }

                    if ((Tipo == "DGD") || (Tipo == "DFA"))//para devolucion y nota de credito
                    {
                        if (bDividir == true)
                        {
                            if (CantxUni != 0)
                                oDocuments.Lines.Quantity = dquantity / CantxUni;
                            else
                                oDocuments.Lines.Quantity = dquantity;
                        }
                        else
                            oDocuments.Lines.Quantity = dquantity;
                    }
                    else
                    oDocuments.Lines.Quantity = dquantity;
                    
                    if (dtlote.Rows.Count > 0)
                        Cambiar_Status_Lote(ref dtlote, false);

                    lRetCode = oDocuments.Add();
                    if (lRetCode != 0)
                    {
                        oCompany.GetLastError(out lErrcode, out sErrmsg);
                        Func.AddLog("Archivo " + sv + " con problemas al crear en SAP, " + sErrmsg);
                        Func.AddLogSap(oCompany, "Archivo " + sv + " con problemas al crear en SAP, " + sErrmsg, sv, dquantity, "", "", "", "", "", "", ItemCode, "", "", "");
                        _return = false;
                        s = "C:\\Documents Ma.xml";
                        oDocuments.SaveXML(s);

                        //s = PathNoOk + "\\" + sv.Substring(PathArchivos.Length, sv.Length - PathArchivos.Length);
                        //if (File.Exists(s))
                        //    File.Delete(s);
                        //File.Move(sv, s);
                    }
                    else
                    {
                        _return = true;
                        if (Tipo == "IMP")
                        {
                            sDocEntry = oCompany.GetNewObjectKey();
                            oDocuments = null;
                            oDocuments = ((SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices));
                            oDocuments2 = ((SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes));
                            if (oDocuments.GetByKey(Convert.ToInt32(sDocEntry)))
                            {
                                oDocuments2.CardCode = CardCode;
                                oDocuments2.DocDate = DocDate;
                                ilinea = 0;
                                ilotelinea = 0;
                                DocEntryOC = oDocuments.Lines.BaseEntry;

                                while (ilinea < oDocuments.Lines.Count)
                                {
                                    if (ilinea > 0)
                                        oDocuments2.Lines.Add();
                                    oDocuments.Lines.SetCurrentLine(ilinea);
                                    oDocuments2.Lines.ItemCode = oDocuments.Lines.ItemCode;
                                    oDocuments2.Lines.BaseType = 18;
                                    oDocuments2.Lines.BaseEntry = oDocuments.DocEntry;
                                    oDocuments2.Lines.BaseLine = oDocuments.Lines.LineNum;
                                    ilotelinea = 0;
                                    dquantity = 0;

                                    foreach (System.Data.DataRow orows in dtSorted.Rows)
                                    {
                                        if (oDocuments.Lines.ItemCode == ((System.String)orows["ItemCode"]).Trim())
                                        {
                                            s = "SELECT ManBtchNum FROM OITM WHERE ItemCode = '{0}'";
                                            s = String.Format(s, ((System.String)orows["ItemCode"]).Trim());
                                            oRecordSet.DoQuery(s);
                                            if (oRecordSet.RecordCount == 0)
                                                ManejaLote = false;
                                            else
                                                if (((System.String)oRecordSet.Fields.Item("ManBtchNum").Value).Trim() == "Y")
                                                    ManejaLote = true;
                                                else
                                                    ManejaLote = false;
                                            if (ManejaLote)
                                            {
                                                if (ilotelinea > 0)
                                                    oDocuments2.Lines.BatchNumbers.Add();
                                                //if (((System.DateTime)orows["ExpDate"]) != null)
                                                if (orows["ExpDate"].ToString() != "")
                                                    oDocuments2.Lines.BatchNumbers.ExpiryDate = ((System.DateTime)orows["ExpDate"]);
                                                oDocuments2.Lines.BatchNumbers.Quantity = ((System.Double)orows["Quantity"]);
                                                oDocuments2.Lines.BatchNumbers.BatchNumber = ((System.String)orows["Lote"]).Trim();
                                                dquantity = dquantity + ((System.Double)orows["Quantity"]);
                                                ilotelinea++;
                                                //guardo lote en dt para actualizar status
                                                s = @"select Status from OBTN where ItemCode = '{0}' and DistNumber = '{1}' and Status = '1'";
                                                s = String.Format(s, ((System.String)orows["ItemCode"]).Trim(), ((System.String)orows["Lote"]));
                                                oRecordSet.DoQuery(s);
                                                if (oRecordSet.RecordCount > 0)
                                                {
                                                    dtrow = dtlote.NewRow();
                                                    dtrow["ItemCode"] = ((System.String)orows["ItemCode"]).Trim();
                                                    dtrow["Lote"] = ((System.String)orows["Lote"]);
                                                    dtrow["Status"] = ((System.String)oRecordSet.Fields.Item("Status").Value).Trim();
                                                    dtlote.Rows.Add(dtrow);
                                                }
                                            }
                                            ilotelinea++;
                                        }
                                    }

                                    /*while (ilotelinea < oDocuments.Lines.BatchNumbers.Count)
                                    {
                                        if (ilotelinea > 0)
                                            oDocuments2.Lines.BatchNumbers.Add();
                                        oDocuments.Lines.BatchNumbers.SetCurrentLine(ilotelinea);
                                        oDocuments2.Lines.BatchNumbers.ExpiryDate = oDocuments.Lines.BatchNumbers.ExpiryDate;
                                        oDocuments2.Lines.BatchNumbers.BatchNumber = oDocuments.Lines.BatchNumbers.BatchNumber;
                                        ilotelinea++;
                                    }*/
                                    ilotelinea = 0;
                                    ilinea++;
                                }
                                if (dtlote.Rows.Count > 0)
                                    Cambiar_Status_Lote(ref dtlote, false);

                                lRetCode = oDocuments2.Add();
                                if (lRetCode != 0)
                                {
                                    oCompany.GetLastError(out lErrcode, out sErrmsg);
                                    Func.AddLog("Archivo " + sv + " con problemas al crear en SAP, Entrada mercancia OP basada en factura reserva, " + sErrmsg);
                                    Func.AddLogSap(oCompany, "Archivo " + sv + " con problemas al crear en SAP, Entrada mercancia OP basada en factura reserva, " + sErrmsg, sv, dquantity, "", "", "", "", "", "", ItemCode, "", "", "");
                                    s = "C:\\Documents OP Ma.xml";
                                    oDocuments2.SaveXML(s);
                                    _return = false;
                                }
                                else
                                {
                                    _return = true;
                                    var oOC = ((SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders));
                                    if (oOC.GetByKey(DocEntryOC))
                                    {
                                        if (oOC.DocumentStatus == BoStatus.bost_Open)
                                        {
                                            lRetCode = oOC.Close();
                                            if (lRetCode != 0)
                                            {
                                                oCompany.GetLastError(out lErrcode, out sErrmsg);
                                                Func.AddLog("Archivo " + sv + " con problemas al cerrar OC " + oDocuments.DocNum.ToString());
                                                Func.AddLogSap(oCompany, "Archivo " + sv + " con problemas al cerrar OC ", sv, dquantity, "", "", "", "", "", "", ItemCode, "", "", "");
                                            }
                                            else
                                            {
                                                Func.AddLog("Archivo " + sv + " OC cerrada por la interfaz");
                                                Func.AddLogSap(oCompany, "Archivo " + sv + " OC cerrada por la interfaz", sv, dquantity, "", "", "", "", "", "", ItemCode, "", "", "");
                                            }
                                        }
                                    }

                                }
                            }
                        }
                        else if (Tipo == "NAC")
                        {
                            var oOC = ((SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders));
                            if (oOC.GetByKey(DocEntryOC))
                            {
                                if (oOC.DocumentStatus == BoStatus.bost_Open)
                                {
                                    lRetCode = oOC.Close();
                                    if (lRetCode != 0)
                                    {
                                        oCompany.GetLastError(out lErrcode, out sErrmsg);
                                        Func.AddLog("Archivo " + sv + " con problemas al cerrar OC " + oDocuments.DocNum.ToString());
                                        Func.AddLogSap(oCompany, "Archivo " + sv + " con problemas al cerrar OC " + oDocuments.DocNum.ToString(), sv, dquantity, "", "", "", "", "", "", ItemCode, "", "", "");
                                    }
                                    else
                                    {
                                        Func.AddLog("Archivo " + sv + " OC cerrada por la interfaz");
                                        Func.AddLogSap(oCompany, "Archivo " + sv + " OC cerrada por la interfaz", sv, dquantity, "", "", "", "", "", "", ItemCode, "", "", "");
                                    }
                                }
                            }
                        }

                        if (dtlote.Rows.Count > 0)
                            Cambiar_Status_Lote(ref dtlote, true);

                        if (_return)
                        {
                            Func.AddLog("Archivo " + sv + " creado satisfactoriamente en SAP");
                            Func.AddLogSap(oCompany, "Archivo " + sv + " creado satisfactoriamente en SAP", sv, dquantity, "", "", "", "", "", "", ItemCode, "", "", "");
                        }
                    }
                }
                return _return;
            }
            catch (Exception we)
            {
                Func.AddLog("Error SVS_Documents - " + we.Message + ", StackTrace " + we.StackTrace);
                return false;
            }
        }

        //crea objetos trabajo con lote, crea factura y Entrega
        public Boolean SLS_Documents(String Tipo, ref System.Data.DataTable dtsls, ref System.Data.DataTable dtSorted, String sv)
        {
            Boolean _return = false;
            TFunctions Func = new TFunctions();
            SAPbobsCOM.Documents oDocuments = null;
            SAPbobsCOM.Documents oDocuments2;
            SAPbobsCOM.StockTransfer oStockTransfer = null;
            Int32 ilinea;
            Int32 ilotelinea;
            String ItemCode;
            Double dquantity;
            Double CantxUni = 0;
            Boolean bDividir = false;
            Int32 lRetCode;
            Int32 lErrcode;
            String sErrmsg;
            String sDocEntry;
            Boolean bGuia = false;
            Boolean bFactura = false;
            Boolean bBoleta = false;
            DataRow dtrow;
            String CardCode = "";
            String WhsCode = "";
            String Carton = "";
            Boolean ManejaLote;

            try
            {

                foreach (System.Data.DataRow orowsls in dtsls.Rows)
                {
                    try
                    {
                        if (((System.String)orowsls["DocEntry"]).Trim() == "0") //para crear transferencias
                        {
                            s = @"select O0.QryGroup5 'Transferncia' from OCRD O0 where O0.CardCode = '{0}'";
                            s = String.Format(s, ((System.String)orowsls["CardCode"]).Trim());
                            oRecordSet.DoQuery(s);

                            if (oRecordSet.RecordCount == 0)
                            {
                                Func.AddLog("Archivo " + sv + " con problemas al crear en SAP, Cliente no se encuentra en la base de datos");
                                _return = false;
                            }
                            else
                            {
                                oStockTransfer = ((SAPbobsCOM.StockTransfer)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer));
                                CardCode = ((System.String)orowsls["CardCode"]).Trim();
                                oStockTransfer.CardCode = CardCode;
                                oStockTransfer.DocDate = DateTime.Now;
                                Carton = ((System.String)orowsls["Carton"]).Trim();
                                oStockTransfer.UserFields.Fields.Item("U_NCarton").Value = Carton;

                                ItemCode = "";
                                dquantity = 0;
                                ilotelinea = 0;
                                ilinea = 0;
                                foreach (System.Data.DataRow orows in dtSorted.Rows)
                                {
                                    if (((System.String)orows["CardCode"]).Trim() == CardCode)
                                    {
                                        if (ItemCode != ((System.String)orows["ItemCode"]).Trim())
                                        {
                                            oStockTransfer.Lines.Quantity = dquantity;

                                            if (ilinea > 0)
                                            {
                                                oStockTransfer.Lines.Add();
                                                ilotelinea = 0;
                                                dquantity = 0;
                                            }
                                            else
                                            {
                                                oStockTransfer.Address = ((System.String)orows["Address"]).Trim();
                                                oStockTransfer.FromWarehouse = "02";
                                                oStockTransfer.ToWarehouse = CardCode.Substring(3, 2);
                                            }

                                            oStockTransfer.Lines.ItemCode = ((System.String)orows["ItemCode"]).Trim();
                                            oStockTransfer.Lines.WarehouseCode = CardCode.Substring(3, 2);
                                            oStockTransfer.Lines.Quantity = dquantity;
                                        }

                                        s = "SELECT ManBtchNum FROM OITM WHERE ItemCode = '{0}'";
                                        s = String.Format(s, ((System.String)orows["ItemCode"]).Trim());
                                        oRecordSet.DoQuery(s);
                                        if (oRecordSet.RecordCount == 0)
                                            ManejaLote = false;
                                        else
                                            if (((System.String)oRecordSet.Fields.Item("ManBtchNum").Value).Trim() == "Y")
                                                ManejaLote = true;
                                            else
                                                ManejaLote = false;

                                        if ((((System.String)orows["Lote"]).Trim() != "") && (ManejaLote))
                                        {
                                            if (ilotelinea > 0)
                                                oStockTransfer.Lines.BatchNumbers.Add();
                                            if (orows["ExpDate"].ToString() != "")
                                                oStockTransfer.Lines.BatchNumbers.ExpiryDate = ((System.DateTime)orows["ExpDate"]);
                                            oStockTransfer.Lines.BatchNumbers.Quantity = ((System.Double)orows["Quantity"]);
                                            oStockTransfer.Lines.BatchNumbers.BatchNumber = ((System.String)orows["Lote"]).Trim();
                                            dquantity = dquantity + ((System.Double)orows["Quantity"]);
                                            ilotelinea++;

                                            //guardo lote en dt para actualizar status
                                            s = @"select Status from OBTN where ItemCode = '{0}' and DistNumber = '{1}' and Status = '1'";
                                            s = String.Format(s, ((System.String)orows["ItemCode"]).Trim(), ((System.String)orows["Lote"]));
                                            oRecordSet.DoQuery(s);
                                            if (oRecordSet.RecordCount > 0)
                                            {
                                                dtrow = dtlote.NewRow();
                                                dtrow["ItemCode"] = ((System.String)orows["ItemCode"]).Trim();
                                                dtrow["Lote"] = ((System.String)orows["Lote"]);
                                                dtrow["Status"] = ((System.String)oRecordSet.Fields.Item("Status").Value).Trim();
                                                dtlote.Rows.Add(dtrow);
                                            }
                                        }
                                        else
                                            dquantity = dquantity + ((System.Double)orows["Quantity"]);

                                        ItemCode = ((System.String)orows["ItemCode"]).Trim();
                                        ilinea++;
                                    }
                                }

                                oStockTransfer.Lines.Quantity = dquantity;

                                if (dtlote.Rows.Count > 0)
                                    Cambiar_Status_Lote(ref dtlote, false);

                                lRetCode = oStockTransfer.Add();
                                if (lRetCode != 0)
                                {
                                    oCompany.GetLastError(out lErrcode, out sErrmsg);
                                    Func.AddLog("Archivo " + sv + " con problemas al crear documentos en SAP, " + sErrmsg);
                                    Func.AddLogSap(oCompany, "Archivo " + sv + " con problemas al crear documentos en SAP, " + sErrmsg, sv, dquantity, CardCode
                                                                  , "", "", "", "", "", ItemCode, "", "LPN", Carton);
                                    _return = false;
                                    s = "C:\\oStockTransfer SLS.xml";
                                    oStockTransfer.SaveXML(s);

                                    //s = PathNoOk + "\\" + sv.Substring(PathArchivos.Length, sv.Length - PathArchivos.Length);
                                    //if (File.Exists(s))
                                    //    File.Delete(s);
                                    //File.Move(sv, s);
                                }
                                else
                                {
                                    _return = true;
                                }
                            }
                        }
                        else
                        {
                            s = @"select O0.QryGroup3 'Guia', O0.QryGroup4 'Factura', O0.QryGroup5 'Boleta', T0.DocDueDate 'FechaEntrega' from ORDR T0 join OCRD O0 on O0.CardCode = T0.CardCode where T0.DocEntry = {0}";
                            s = String.Format(s, ((System.String)orowsls["DocEntry"]).Trim());
                            oRecordSet.DoQuery(s);

                            if (oRecordSet.RecordCount == 0)
                            {
                                Func.AddLog("Archivo " + sv + " con problemas al crear en SAP, Cliente no se encuentra en la base de datos");
                                _return = false;
                            }
                            else
                            {
                                if (((System.String)oRecordSet.Fields.Item("Guia").Value).Trim() == "Y")
                                    bGuia = true;
                                else
                                    bGuia = false;

                                if (((System.String)oRecordSet.Fields.Item("Factura").Value).Trim() == "Y")
                                    bFactura = true;
                                else
                                    bFactura = false;

                                if (((System.String)oRecordSet.Fields.Item("Boleta").Value).Trim() == "Y")
                                    bBoleta = true;
                                else
                                    bBoleta = false;


                                if (bGuia)
                                    oDocuments = ((SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes));
                                else
                                    oDocuments = ((SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices));


                                if (oDocuments != null)
                                {
                                    CardCode = ((System.String)orowsls["CardCode"]).Trim();
                                    oDocuments.CardCode = CardCode;
                                    string cnbCodigo = CardCode.Substring(0, 3);
                                    if (cnbCodigo.CompareTo("CNB") == 0)  //si cliente en CNB
                                    {
                                        DateTime dateaux = (System.DateTime)oRecordSet.Fields.Item("FechaEntrega").Value;//sacar fecha de la OV
                                        oDocuments.DocDate = dateaux;
                                    }
                                    else
                                    {
                                        oDocuments.DocDate = DateTime.Now;
                                    }

                                    Carton = ((System.String)orowsls["Carton"]).Trim();
                                    oDocuments.UserFields.Fields.Item("U_NCarton").Value = Carton;

                                    if ((bBoleta) && (!bGuia))
                                        oDocuments.DocumentSubType = BoDocumentSubType.bod_Bill;

                                    ItemCode = "";
                                    dquantity = 0;
                                    ilotelinea = 0;
                                    ilinea = 0;
                                    foreach (System.Data.DataRow orows in dtSorted.Rows)
                                    {
                                        if (Convert.ToString((System.Int32)orows["DocEntry"]) == ((System.String)orowsls["DocEntry"]).Trim())
                                        {
                                            if (ItemCode != ((System.String)orows["ItemCode"]).Trim())
                                            {
                                                if (bDividir == true)
                                                {
                                                    if (CantxUni != 0)
                                                        oDocuments.Lines.Quantity = dquantity / CantxUni;
                                                    else
                                                        oDocuments.Lines.Quantity = dquantity;
                                                }
                                                else
                                                    oDocuments.Lines.Quantity = dquantity;

                                                if (ilinea > 0)
                                                {
                                                    oDocuments.Lines.Add();
                                                    ilotelinea = 0;
                                                    dquantity = 0;
                                                }
                                                else
                                                    oDocuments.Address = ((System.String)orows["Address"]).Trim();

                                                oDocuments.Lines.ItemCode = ((System.String)orows["ItemCode"]).Trim();
                                                //oDocuments.Lines.WarehouseCode = ;

                                                oDocuments.Lines.Quantity = dquantity;
                                                s = "select T1.UseBaseUn, O0.NumInSale, T1.WhsCode from RDR1 T1 join OITM O0 on O0.ItemCode = T1.ItemCode where T1.DocEntry = {0} and T1.LineNum = {1}";
                                                s = String.Format(s, ((System.Int32)orows["DocEntry"]), ((System.Int32)orows["LineNum"]));
                                                oRecordSet.DoQuery(s);
                                                WhsCode = ((System.String)oRecordSet.Fields.Item("WhsCode").Value).Trim();
                                                if (((System.String)oRecordSet.Fields.Item("UseBaseUn").Value).Trim() == "N")
                                                {
                                                    CantxUni = (System.Double)(oRecordSet.Fields.Item("NumInSale").Value);
                                                    bDividir = true;
                                                    if (CantxUni != 0)
                                                        oDocuments.Lines.Quantity = dquantity / CantxUni;
                                                    else
                                                        oDocuments.Lines.Quantity = dquantity;
                                                }
                                                else
                                                    oDocuments.Lines.Quantity = dquantity;


                                                oDocuments.Lines.BaseType = ((System.Int32)orows["ObjType"]);
                                                oDocuments.Lines.BaseEntry = ((System.Int32)orows["DocEntry"]);
                                                oDocuments.Lines.BaseLine = ((System.Int32)orows["LineNum"]);
                                            }

                                            s = "SELECT ManBtchNum FROM OITM WHERE ItemCode = '{0}'";
                                            s = String.Format(s, ((System.String)orows["ItemCode"]).Trim());
                                            oRecordSet.DoQuery(s);
                                            if (oRecordSet.RecordCount == 0)
                                                ManejaLote = false;
                                            else
                                                if (((System.String)oRecordSet.Fields.Item("ManBtchNum").Value).Trim() == "Y")
                                                    ManejaLote = true;
                                                else
                                                    ManejaLote = false;

                                            if ((((System.String)orows["Lote"]).Trim() != "") && (ManejaLote))
                                            {
                                                //consultar por stock del lote en SAP
                                                s = @"select T1.WhsCode, T1.Quantity, T0.BatchNum
                                                          from OIBT T0 
                                                          JOIN OBTQ T1 ON T1.ItemCode = T0.ItemCode 
                                                                      AND T1.SysNumber = T0.SysNumber
			                                                          AND T1.WhsCode = T0.WhsCode
                                                         WHERE T0.ItemCode = '{0}'
                                                           AND T0.BatchNum = '{1}'
                                                           AND T1.WhsCode = '{2}'";
                                                s = String.Format(s, oDocuments.Lines.ItemCode.ToString().Trim(), ((System.String)orows["Lote"]).Trim(), WhsCode);

                                                oRecordSet.DoQuery(s);
                                                if (oRecordSet.RecordCount == 0)
                                                {
                                                    Func.AddLog("Archivo " + sv + " con problemas al crear en SAP, no encuentra articulo con lote indicado - " + oDocuments.Lines.ItemCode.ToString().Trim() + " - " + WhsCode + " - " + ((System.String)orows["Lote"]).Trim());
                                                    return false;
                                                }
                                                else
                                                {
                                                    s = ((System.Double)oRecordSet.Fields.Item("Quantity").Value).ToString();
                                                    //s = ((System.Double)orows["Quantity"]).ToString();
                                                    if ((((System.Double)orows["Quantity"]) > Convert.ToDouble(s, _nf)) || (Convert.ToDouble(s, _nf) <= 0))
                                                    {
                                                        Func.AddLog("Archivo " + sv + " con problemas al crear en SAP, stock insuficiente para articulo con lote indicado - " + oDocuments.Lines.ItemCode.ToString().Trim() + " - " + WhsCode + " - " + ((System.String)orows["Lote"]).Trim());
                                                        return false;
                                                    }
                                                }


                                                if (ilotelinea > 0)
                                                    oDocuments.Lines.BatchNumbers.Add();
                                                if (orows["ExpDate"].ToString() != "")
                                                    oDocuments.Lines.BatchNumbers.ExpiryDate = ((System.DateTime)orows["ExpDate"]);
                                                oDocuments.Lines.BatchNumbers.Quantity = ((System.Double)orows["Quantity"]);
                                                oDocuments.Lines.BatchNumbers.BatchNumber = ((System.String)orows["Lote"]).Trim();
                                                dquantity = dquantity + ((System.Double)orows["Quantity"]);
                                                ilotelinea++;

                                                //guardo lote en dt para actualizar status
                                                s = @"select Status from OBTN where ItemCode = '{0}' and DistNumber = '{1}' and Status = '1'";
                                                s = String.Format(s, ((System.String)orows["ItemCode"]).Trim(), ((System.String)orows["Lote"]));
                                                oRecordSet.DoQuery(s);
                                                if (oRecordSet.RecordCount > 0)
                                                {
                                                    dtrow = dtlote.NewRow();
                                                    dtrow["ItemCode"] = ((System.String)orows["ItemCode"]).Trim();
                                                    dtrow["Lote"] = ((System.String)orows["Lote"]);
                                                    dtrow["Status"] = ((System.String)oRecordSet.Fields.Item("Status").Value).Trim();
                                                    dtlote.Rows.Add(dtrow);
                                                }
                                            }
                                            else
                                                dquantity = dquantity + ((System.Double)orows["Quantity"]);

                                            ItemCode = ((System.String)orows["ItemCode"]).Trim();
                                            ilinea++;
                                        }
                                    }

                                    if (bDividir == true)
                                    {
                                        if (CantxUni != 0)
                                            oDocuments.Lines.Quantity = dquantity / CantxUni;
                                        else
                                            oDocuments.Lines.Quantity = dquantity;
                                    }
                                    else
                                        oDocuments.Lines.Quantity = dquantity;

                                    if (dtlote.Rows.Count > 0)
                                        Cambiar_Status_Lote(ref dtlote, false);

                                    lRetCode = oDocuments.Add();
                                    if (lRetCode != 0)
                                    {
                                        oCompany.GetLastError(out lErrcode, out sErrmsg);
                                        Func.AddLog("Archivo " + sv + " con problemas al crear en SAP, " + sErrmsg);
                                        Func.AddLogSap(oCompany, "Archivo " + sv + " con problemas al crear en SAP, " + sErrmsg, sv, dquantity, CardCode
                                                                  , "", "", "", "", "", ItemCode, "", "LPN", Carton);
                                        _return = false;
                                        s = "C:\\Documents SLS.xml";
                                        oDocuments.SaveXML(s);

                                        //s = PathNoOk + "\\" + sv.Substring(PathArchivos.Length, sv.Length - PathArchivos.Length);
                                        //if (File.Exists(s))
                                        //    File.Delete(s);
                                        //File.Move(sv, s);
                                    }
                                    else
                                    {
                                        _return = true;

                                        string tipoDocSap = "";

                                        if (((bGuia) && (!bFactura)) || ((bGuia) && (bFactura)))
                                            tipoDocSap = "Guia";
                                        else
                                            if ((!bGuia) && (bFactura))
                                                tipoDocSap = "Factura";

                                        if (tipoDocSap != "")
                                        {
                                            sDocEntry = oCompany.GetNewObjectKey();

                                            foreach (System.Data.DataRow orows in dtSorted.Rows)
                                            {
                                                //preguntar por carton sea igual al de la factura o guia 
                                                string codigo = ((System.String)orows["ItemCode"]).Trim();
                                                int cantidad = Convert.ToInt32((System.Double)orows["Quantity"]);
                                                string lote = ((System.String)orows["Lote"]).Trim();
                                                DateTime timeExp = (System.DateTime)orows["ExpDate"];
                                                int fkfactura = (Int32.Parse(sDocEntry));
                                                string cartonAux = ((System.String)orows["Carton"]).Trim();
                                                if (Carton == cartonAux)
                                                    Func.insertarLotesBDAux(codigo, cantidad, lote, timeExp, tipoDocSap, fkfactura, cartonAux);
                                            }

                                            //si creo guia o factura, Busco de la Orden de Venta los campos de direccion de facturación y Destino, y lo asigno ya que no los toma por defecto del documento base
                                            if (tipoDocSap == "Guia")
                                                oDocuments = ((SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes));
                                            else
                                                oDocuments = ((SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices));

                                            if (oDocuments.GetByKey(Convert.ToInt32(sDocEntry)))
                                            {
                                                s = @"SELECT ISNULL(T0.StreetB,'') as 'StreetB', ISNULL(T0.StreetNoB,'') as 'StreetNoB', ISNULL(T0.BlockB,'') as 'BlockB' ,
		                                                    ISNULL(T0.CityB,'') as 'CityB' , ISNULL(T0.ZipCodeB,'') as 'ZipCodeB', ISNULL(T0.CountyB,'') as 'CountyB' ,
		                                                    ISNULL(T0.StateB,'') as 'StateB', ISNULL(T0.CountryB,'') as 'CountryB',
		                                                    ISNULL(T0.StreetS,'') as 'StreetS',  ISNULL(T0.StreetNoS,'') as 'StreetNoS', ISNULL(T0.BlockS,'') as 'BlockS' ,
		                                                    ISNULL(T0.CityS,'') as 'CityS' , ISNULL(T0.ZipCodeS,'') as 'ZipCodeS', ISNULL(T0.CountyS,'') as 'CountyS' ,
		                                                    ISNULL(T0.StateS,'') as 'StateS', ISNULL(T0.CountryS,'') as 'CountryS'
                                                      FROM  RDR12 T0  WHERE T0.DocEntry =  {0}";

                                                s = String.Format(s, ((System.String)orowsls["DocEntry"]).Trim());
                                                oRecordSet.DoQuery(s);

                                                if (oRecordSet.RecordCount == 0)
                                                {
                                                    Func.AddLog("Error al obtener direccion de facturacion de la Orden de Venta: " + ((System.String)orowsls["DocEntry"]).Trim());
                                                }
                                                else
                                                {
                                                    AddressExtension adreesExtension1;
                                                    adreesExtension1 = (AddressExtension)oDocuments.AddressExtension;
                                                    //Facturacion
                                                    adreesExtension1.BillToStreet = ((System.String)oRecordSet.Fields.Item("StreetB").Value).Trim();
                                                    adreesExtension1.BillToStreetNo = ((System.String)oRecordSet.Fields.Item("StreetNoB").Value).Trim();
                                                    adreesExtension1.BillToBlock = ((System.String)oRecordSet.Fields.Item("BlockB").Value).Trim();
                                                    adreesExtension1.BillToCity = ((System.String)oRecordSet.Fields.Item("CityB").Value).Trim();
                                                    adreesExtension1.BillToZipCode = ((System.String)oRecordSet.Fields.Item("ZipCodeB").Value).Trim();
                                                    adreesExtension1.BillToCounty = ((System.String)oRecordSet.Fields.Item("CountyB").Value).Trim();
                                                    adreesExtension1.BillToState = ((System.String)oRecordSet.Fields.Item("StateB").Value).Trim();
                                                    adreesExtension1.BillToCountry = ((System.String)oRecordSet.Fields.Item("CountryB").Value).Trim();
                                                    //Despacho  *no me permite modificar una guia basada en OV
                                               //     adreesExtension1.ShipToStreet = ((System.String)oRecordSet.Fields.Item("StreetS").Value).Trim();
                                               //     adreesExtension1.ShipToStreetNo = ((System.String)oRecordSet.Fields.Item("StreetNoS").Value).Trim();
                                               //     adreesExtension1.ShipToBlock = ((System.String)oRecordSet.Fields.Item("BlockS").Value).Trim();
                                               //     adreesExtension1.ShipToCity = ((System.String)oRecordSet.Fields.Item("CityS").Value).Trim();
                                               //     adreesExtension1.ShipToZipCode = ((System.String)oRecordSet.Fields.Item("ZipCodeS").Value).Trim();
                                               //     adreesExtension1.ShipToCounty = ((System.String)oRecordSet.Fields.Item("CountyS").Value).Trim();
                                               //     adreesExtension1.ShipToState = ((System.String)oRecordSet.Fields.Item("StateS").Value).Trim();
                                               //     adreesExtension1.ShipToCountry = ((System.String)oRecordSet.Fields.Item("CountryS").Value).Trim();

                                                    lRetCode = oDocuments.Update();

                                                    if (lRetCode != 0)
                                                    {
                                                        oCompany.GetLastError(out lErrcode, out sErrmsg);
                                                        Func.AddLog("Error en Actualizacion de direccion Documento: " + tipoDocSap + " Mensaje: " + sErrmsg + " " + lErrcode);
                                                    }
                                                }
                                            }

                                        }

                                        if ((bGuia) && (bFactura))  //Escenario de crear factura basado en una entrega
                                        {
                                            sDocEntry = oCompany.GetNewObjectKey();
                                            oDocuments = null;
                                            oDocuments = ((SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes));
                                            oDocuments2 = ((SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices));
                                            if (oDocuments.GetByKey(Convert.ToInt32(sDocEntry)))
                                            {
                                                oDocuments2.CardCode = CardCode;
                                                oDocuments2.DocDate = oDocuments.DocDate;
                                                oDocuments2.UserFields.Fields.Item("U_NCarton").Value = oDocuments.UserFields.Fields.Item("U_NCarton").Value;
                                                ilinea = 0;
                                                ilotelinea = 0;

                                                oDocuments2.AddressExtension.BillToBlock = oDocuments.AddressExtension.BillToBlock;
                                                oDocuments2.AddressExtension.BillToBuilding = oDocuments.AddressExtension.BillToBuilding;
                                                oDocuments2.AddressExtension.BillToCity = oDocuments.AddressExtension.BillToCity;
                                                oDocuments2.AddressExtension.BillToCountry = oDocuments.AddressExtension.BillToCountry;
                                                oDocuments2.AddressExtension.BillToCounty = oDocuments.AddressExtension.BillToCounty;
                                                oDocuments2.AddressExtension.BillToState = oDocuments.AddressExtension.BillToState;
                                                oDocuments2.AddressExtension.BillToStreet = oDocuments.AddressExtension.BillToStreet;
                                                oDocuments2.AddressExtension.BillToStreetNo = oDocuments.AddressExtension.BillToStreetNo;
                                                oDocuments2.AddressExtension.BillToZipCode = oDocuments.AddressExtension.BillToZipCode;
                                                oDocuments2.AddressExtension.BillToAddressType = oDocuments.AddressExtension.BillToAddressType;

                                                oDocuments2.AddressExtension.ShipToBlock = oDocuments.AddressExtension.ShipToBlock;
                                                oDocuments2.AddressExtension.ShipToBuilding = oDocuments.AddressExtension.ShipToBuilding;
                                                oDocuments2.AddressExtension.ShipToCity = oDocuments.AddressExtension.ShipToCity;
                                                oDocuments2.AddressExtension.ShipToCountry = oDocuments.AddressExtension.ShipToCountry;
                                                oDocuments2.AddressExtension.ShipToState = oDocuments.AddressExtension.ShipToState;
                                                oDocuments2.AddressExtension.ShipToStreet = oDocuments.AddressExtension.ShipToStreet;
                                                oDocuments2.AddressExtension.ShipToStreetNo = oDocuments.AddressExtension.ShipToStreetNo;
                                                oDocuments2.AddressExtension.ShipToZipCode = oDocuments.AddressExtension.ShipToZipCode;
                                                string comuna =  oDocuments.AddressExtension.ShipToCounty;
                                                oDocuments2.AddressExtension.ShipToCounty = oDocuments.AddressExtension.ShipToCounty;

                                                while (ilinea < oDocuments.Lines.Count)
                                                {
                                                    if (ilinea > 0)
                                                        oDocuments2.Lines.Add();
                                                    oDocuments.Lines.SetCurrentLine(ilinea);
                                                    oDocuments2.Lines.ItemCode = oDocuments.Lines.ItemCode;
                                                    oDocuments2.Lines.BaseType = 15;
                                                    oDocuments2.Lines.BaseEntry = oDocuments.DocEntry;
                                                    oDocuments2.Lines.BaseLine = oDocuments.Lines.LineNum;

                                                    ////no deberia necesitar lote debido que se encuentra en la guia
                                                    //while (ilotelinea < oDocuments.Lines.BatchNumbers.Count)
                                                    //{
                                                    //    if (ilotelinea > 0)
                                                    //        oDocuments2.Lines.BatchNumbers.Add();
                                                    //    oDocuments.Lines.BatchNumbers.SetCurrentLine(ilotelinea);
                                                    //    oDocuments2.Lines.BatchNumbers.ExpiryDate = oDocuments.Lines.BatchNumbers.ExpiryDate;
                                                    //    oDocuments2.Lines.BatchNumbers.BatchNumber = oDocuments.Lines.BatchNumbers.BatchNumber;
                                                    //    ilotelinea++;
                                                    //}
                                                    ilotelinea = 0;
                                                    ilinea++;
                                                }

                                                lRetCode = oDocuments2.Add();
                                                if (lRetCode != 0)
                                                {
                                                    oCompany.GetLastError(out lErrcode, out sErrmsg);
                                                    Func.AddLog("Archivo " + sv + " con problemas al crear en SAP, factura en base Entrega, " + sErrmsg);
                                                    Func.AddLogSap(oCompany, "Archivo " + sv + " con problemas al crear en SAP, factura en base Entrega, " + sErrmsg, sv, dquantity, CardCode
                                                                  , "", "", "", "", "", ItemCode, "", "LPN", Carton);
                                                    _return = false;
                                                }
                                                else
                                                {
                                                    sDocEntry = oCompany.GetNewObjectKey();

                                                    foreach (System.Data.DataRow orows in dtSorted.Rows)
                                                    {
                                                        string codigo = ((System.String)orows["ItemCode"]).Trim();
                                                        int cantidad = Convert.ToInt32((System.Double)orows["Quantity"]);
                                                        string lote = ((System.String)orows["Lote"]).Trim();
                                                        DateTime timeExp = (System.DateTime)orows["ExpDate"];
                                                        int fkfactura = (Int32.Parse(sDocEntry));

                                                        string cartonAux = ((System.String)orows["Carton"]).Trim();
                                                        if (Carton == cartonAux)
                                                            Func.insertarLotesBDAux(codigo, cantidad, lote, timeExp, "Factura", fkfactura, cartonAux);
                                                    }

                                                    //actualizacion de direccion basado en la nota de venta de la factura 
                                                    oDocuments = ((SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices));
                                                    if (oDocuments.GetByKey(Convert.ToInt32(sDocEntry)))
                                                    {
                                                        s = @"SELECT ISNULL(T0.StreetB,'') as 'StreetB', ISNULL(T0.StreetNoB,'') as 'StreetNoB', ISNULL(T0.BlockB,'') as 'BlockB' ,
		                                                             ISNULL(T0.CityB,'') as 'CityB' , ISNULL(T0.ZipCodeB,'') as 'ZipCodeB', ISNULL(T0.CountyB,'') as 'CountyB' ,
		                                                             ISNULL(T0.StateB,'') as 'StateB', ISNULL(T0.CountryB,'') as 'CountryB',
		                                                             ISNULL(T0.StreetS,'') as 'StreetS',  ISNULL(T0.StreetNoS,'') as 'StreetNoS', ISNULL(T0.BlockS,'') as 'BlockS' ,
		                                                             ISNULL(T0.CityS,'') as 'CityS' , ISNULL(T0.ZipCodeS,'') as 'ZipCodeS', ISNULL(T0.CountyS,'') as 'CountyS' ,
		                                                             ISNULL(T0.StateS,'') as 'StateS', ISNULL(T0.CountryS,'') as 'CountryS'
                                                            FROM  RDR12 T0  WHERE T0.DocEntry =  {0}";
                                                        
                                                        s = String.Format(s, ((System.String)orowsls["DocEntry"]).Trim());
                                                        oRecordSet.DoQuery(s);

                                                        if (oRecordSet.RecordCount == 0)
                                                        {
                                                            Func.AddLog("Error al obtener direccion de facturacion de la Orden de Venta: " + ((System.String)orowsls["DocEntry"]).Trim());
                                                        }
                                                        else
                                                        {
                                                            AddressExtension adreesExtension1;
                                                            adreesExtension1 = (AddressExtension)oDocuments.AddressExtension;

                                                            //Facturacion
                                                            adreesExtension1.BillToStreet = ((System.String)oRecordSet.Fields.Item("StreetB").Value).Trim();
                                                            adreesExtension1.BillToStreetNo = ((System.String)oRecordSet.Fields.Item("StreetNoB").Value).Trim();
                                                            adreesExtension1.BillToBlock = ((System.String)oRecordSet.Fields.Item("BlockB").Value).Trim();
                                                            adreesExtension1.BillToCity = ((System.String)oRecordSet.Fields.Item("CityB").Value).Trim();
                                                            adreesExtension1.BillToZipCode = ((System.String)oRecordSet.Fields.Item("ZipCodeB").Value).Trim();
                                                            adreesExtension1.BillToCounty = ((System.String)oRecordSet.Fields.Item("CountyB").Value).Trim();
                                                            adreesExtension1.BillToState = ((System.String)oRecordSet.Fields.Item("StateB").Value).Trim();
                                                            adreesExtension1.BillToCountry = ((System.String)oRecordSet.Fields.Item("CountryB").Value).Trim();
                                                            
                                                            //Despacho
                                                            adreesExtension1.ShipToStreet = ((System.String)oRecordSet.Fields.Item("StreetS").Value).Trim();
                                                            adreesExtension1.ShipToStreetNo = ((System.String)oRecordSet.Fields.Item("StreetNoS").Value).Trim();
                                                            adreesExtension1.ShipToBlock = ((System.String)oRecordSet.Fields.Item("BlockS").Value).Trim();
                                                            adreesExtension1.ShipToCity = ((System.String)oRecordSet.Fields.Item("CityS").Value).Trim();
                                                            adreesExtension1.ShipToZipCode = ((System.String)oRecordSet.Fields.Item("ZipCodeS").Value).Trim();
                                                            adreesExtension1.ShipToCounty = ((System.String)oRecordSet.Fields.Item("CountyS").Value).Trim();
                                                            adreesExtension1.ShipToState = ((System.String)oRecordSet.Fields.Item("StateS").Value).Trim();
                                                            adreesExtension1.ShipToCountry = ((System.String)oRecordSet.Fields.Item("CountryS").Value).Trim();

                                                            lRetCode = oDocuments.Update();

                                                            if (lRetCode != 0)
                                                            {
                                                                oCompany.GetLastError(out lErrcode, out sErrmsg);
                                                                Func.AddLog("Error en Actualizacion de direccion Documento: Factura Basada en Entrega  Mensaje: " + sErrmsg + " " + lErrcode);
                                                            }
                                                        }
                                                    }
                                                    _return = true;
                                                }
                                            }
                                        }
                                            if (dtlote.Rows.Count > 0)
                                                Cambiar_Status_Lote(ref dtlote, true);

                                            if (_return)
                                            {
                                                Func.AddLog("Archivo " + sv + " creado satisfactoriamente en SAP");
                                                Func.AddLogSap(oCompany, "Archivo " + sv + " creado satisfactoriamente en SAP", sv, dquantity, CardCode
                                                              , "", "", "", "", "", ItemCode, "", "LPN", Carton);
                                            }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception n)
                    {
                        Func.AddLog("Error archivo " + sv + ", al procesar DocEntry " + ((System.String)orowsls["DocEntry"]).Trim() + " - " + n.Message + ", TRACE " + n.StackTrace);
                    }
                    finally
                    {
                        oDocuments = null;
                        oDocuments2 = null;
                        oStockTransfer = null;
                    }

                }

                return _return;
            }
            catch (Exception we)
            {
                Func.AddLog("Error SLS_Documents - " + we.Message + ", StackTrace " + we.StackTrace);
                return false;
            }
        }

        //crea objetos de Transferencia de Inventario
        public Boolean IHT_Documents(ref System.Data.DataTable dt, String sv)
        {
            TFunctions Func = new TFunctions();
            SAPbobsCOM.Documents oDocuments = null;
            SAPbobsCOM.StockTransfer oTransferStock;
            Int32 lRetCode;
            Int32 lErrcode;
            String sErrmsg;
            Boolean bGenerarAjuste;
            Boolean bConLote;

            try
            {
                foreach (System.Data.DataRow orows in dt.Rows)
                {
                    try
                    {
                        s = "select ManBtchNum from oitm where ItemCode = '{0}'";
                        s = String.Format(s, (System.String)orows["ItemCode"]);
                        oRecordSet.DoQuery(s);
                        if (oRecordSet.RecordCount == 0)
                            throw new Exception("Articulo no se ha encontrado en la base de datos - " + (System.String)orows["ItemCode"]);

                        if (((System.String)oRecordSet.Fields.Item("ManBtchNum").Value).Trim() == "N")
                            bConLote = false;
                        else
                            bConLote = true;

                        if (((System.String)orows["Code"] == "17") && (((System.String)orows["Razon"] == "SP") || ((System.String)orows["Razon"] == "MLPN")))
                        {

                            if ((System.Double)orows["Cantidad"] > 0)
                            {
                                bGenerarAjuste = false;

                                if (bConLote == false)
                                {
                                    s = @"select OnHand 'Quantity' from OITW where ItemCode = '{0}' and WhsCode = '{1}' and OnHand > 0";
                                    s = String.Format(s, (System.String)orows["ItemCode"], "03");
                                }
                                else
                                {
                                    s = @"select T1.Quantity 
                                        from OBTN T0 
                                        JOIN OBTQ T1 ON T1.ItemCode = T0.ItemCode 
                                                    and T1.SysNumber = T0.SysNumber 
                                       where T0.ItemCode = '{0}' 
                                         and T1.WhsCode = '{2}' 
                                         and T0.DistNumber = '{1}'
                                         and T1.Quantity > 0";
                                    s = String.Format(s, (System.String)orows["ItemCode"], (System.String)orows["Ref1"], "03");
                                }
                                oRecordSet.DoQuery(s);
                                if (oRecordSet.RecordCount > 0)
                                {
                                    if (((System.Double)oRecordSet.Fields.Item("Quantity").Value) >= (System.Double)orows["Cantidad"])
                                        bGenerarAjuste = true;

                                    if (bGenerarAjuste)
                                    {
                                        oTransferStock = ((SAPbobsCOM.StockTransfer)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer));
                                        oTransferStock.DocDate = DateTime.Now;
                                        oTransferStock.FromWarehouse = "03";
                                        if ((System.String)orows["CodBloqueo"] == "")
                                            oTransferStock.ToWarehouse = "02";
                                        else if ((System.String)orows["CodBloqueo"] == "B07")
                                            oTransferStock.ToWarehouse = "07";
                                        else if ((System.String)orows["CodBloqueo"] == "B99")
                                            oTransferStock.ToWarehouse = "99";
                                        else if ((System.String)orows["CodBloqueo"] == "CENABAST")
                                            oTransferStock.ToWarehouse = "06";
                                        else if ((System.String)orows["CodBloqueo"] == "VTA")
                                            oTransferStock.ToWarehouse = "08";
                                        else if ((System.String)orows["CodBloqueo"] == "LOTE")
                                            oTransferStock.ToWarehouse = "08";
                                        else if ((System.String)orows["CodBloqueo"] == "PRILOGIC")
                                            oTransferStock.ToWarehouse = "09";
                                        oTransferStock.Lines.ItemCode = (System.String)orows["ItemCode"];
                                        oTransferStock.Lines.Quantity = (System.Double)orows["Cantidad"];

                                        if (bConLote == true)
                                        {
                                            oTransferStock.Lines.BatchNumbers.Quantity = (System.Double)orows["Cantidad"];
                                            oTransferStock.Lines.BatchNumbers.BatchNumber = (System.String)orows["Ref1"];
                                        }

                                        lRetCode = oTransferStock.Add();
                                        if (lRetCode != 0)
                                        {
                                            oCompany.GetLastError(out lErrcode, out sErrmsg);
                                            Func.AddLog("Error generar transferencia para " + sv + " - " + sErrmsg);
                                            Func.AddLogSap(oCompany, "Error generar transferencia para " + sv + " - " + sErrmsg, sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                           , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", "", "", (System.String)orows["ItemCode"], "", "", "");
                                        }
                                        else
                                        {
                                            Func.AddLog("Transferencia generado para " + sv + " -  tipo 17");
                                            Func.AddLogSap(oCompany, "Transferencia generado para " + sv + "tipo 17", sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"] 
                                                           ,(System.String)orows["Razon"], (System.String)orows["Ref1"], "", "", "", (System.String)orows["ItemCode"], "", "", "");
                                        }
                                    }

                                }
                                else
                                {
                                    oDocuments = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(BoObjectTypes.oInventoryGenEntry);
                                    oDocuments.DocDate = DateTime.Now;
                                    oDocuments.Lines.ItemCode = (System.String)orows["ItemCode"];
                                    oDocuments.Lines.Quantity = (System.Double)orows["Cantidad"];
                                    if ((System.String)orows["CodBloqueo"] == "")
                                        oDocuments.Lines.WarehouseCode = "02";
                                    else if ((System.String)orows["CodBloqueo"] == "B07")
                                        oDocuments.Lines.WarehouseCode = "07";
                                    else if ((System.String)orows["CodBloqueo"] == "B99")
                                        oDocuments.Lines.WarehouseCode = "99";
                                    else if ((System.String)orows["CodBloqueo"] == "CENABAST")
                                        oDocuments.Lines.WarehouseCode = "06";
                                    else if ((System.String)orows["CodBloqueo"] == "VTA")
                                        oDocuments.Lines.WarehouseCode = "08";
                                    else if ((System.String)orows["CodBloqueo"] == "LOTE")
                                        oDocuments.Lines.WarehouseCode = "08";
                                    else if ((System.String)orows["CodBloqueo"] == "PRILOGIC")
                                        oDocuments.Lines.WarehouseCode = "09";

                                    if (bConLote == true)
                                    {
                                        oDocuments.Lines.BatchNumbers.Quantity = (System.Double)orows["Cantidad"];
                                        oDocuments.Lines.BatchNumbers.BatchNumber = (System.String)orows["Ref1"];
                                    }

                                    lRetCode = oDocuments.Add();
                                    if (lRetCode != 0)
                                    {
                                        oCompany.GetLastError(out lErrcode, out sErrmsg);
                                        Func.AddLog("Error generar ajuste para " + sv + " - " + sErrmsg);
                                        Func.AddLogSap(oCompany, "Error generar ajuste para " + sv + " - " + sErrmsg, sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                       , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", "", "", (System.String)orows["ItemCode"], "", "", "");
                                    }
                                    else
                                    {
                                        Func.AddLog("Ajuste generado para " + sv + " -  tipo 17");
                                        Func.AddLogSap(oCompany, "Ajuste generado para " + sv + " -  tipo 17", sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                       ,(System.String)orows["Razon"], (System.String)orows["Ref1"], "", "", "", (System.String)orows["ItemCode"], "", "", "");
                                    }
                                }
                            }
                            else
                            {
                                oTransferStock = ((SAPbobsCOM.StockTransfer)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer));
                                oTransferStock.DocDate = DateTime.Now;
                                if ((System.String)orows["CodBloqueo"] == "")
                                    oTransferStock.FromWarehouse = "02";
                                else if ((System.String)orows["CodBloqueo"] == "B07")
                                    oTransferStock.FromWarehouse = "07";
                                else if ((System.String)orows["CodBloqueo"] == "B99")
                                    oTransferStock.FromWarehouse = "99";
                                else if ((System.String)orows["CodBloqueo"] == "CENABAST")
                                    oTransferStock.FromWarehouse = "06";
                                else if ((System.String)orows["CodBloqueo"] == "VTA")
                                    oTransferStock.FromWarehouse = "08";
                                else if ((System.String)orows["CodBloqueo"] == "LOTE")
                                    oTransferStock.FromWarehouse = "08";
                                else if ((System.String)orows["CodBloqueo"] == "PRILOGIC")
                                    oTransferStock.FromWarehouse = "09";

                                oTransferStock.ToWarehouse = "03";
                                oTransferStock.Lines.ItemCode = (System.String)orows["ItemCode"];
                                oTransferStock.Lines.Quantity = (System.Double)orows["Cantidad"] * -1;

                                if (bConLote == true)
                                {
                                    oTransferStock.Lines.BatchNumbers.Quantity = (System.Double)orows["Cantidad"] * -1;
                                    oTransferStock.Lines.BatchNumbers.BatchNumber = (System.String)orows["Ref1"];
                                }

                                lRetCode = oTransferStock.Add();
                                if (lRetCode != 0)
                                {
                                    oCompany.GetLastError(out lErrcode, out sErrmsg);
                                    Func.AddLog("Error generar transferencia para " + sv + " - " + sErrmsg);
                                    Func.AddLogSap(oCompany, "Error generar transferencia para " + sv + " - " + sErrmsg, sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                   , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", "", "", (System.String)orows["ItemCode"], "", "", "");
                                }
                                else
                                {
                                    Func.AddLog("transferencia generado para " + sv + " -  tipo 17");
                                    Func.AddLogSap(oCompany, "Transferencia generado para " + sv + "tipo 17", sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                   ,(System.String)orows["Razon"], (System.String)orows["Ref1"], "", "", "", (System.String)orows["ItemCode"], "", "", "");
                                }
                            }
                        }
                        else if (((System.String)orows["Code"] == "23") && ((System.String)orows["CodBloqueo"] != "PP") && ((System.String)orows["Razon"] == "") &&
                            (((System.String)orows["Ref2"] == "B07") || ((System.String)orows["Ref2"] == "B99") || ((System.String)orows["Ref2"] == "PA") || ((System.String)orows["Ref2"] == "CENABAST" )
                            || ((System.String)orows["Ref2"] == "VTA") || ((System.String)orows["Ref2"] == "LOTE") || ((System.String)orows["Ref2"] == "PRILOGIC")))
                        {
                            if ((System.Double)orows["Cantidad"] > 0)
                            {

                                oTransferStock = ((SAPbobsCOM.StockTransfer)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer));
                                oTransferStock.DocDate = DateTime.Now;
                                oTransferStock.FromWarehouse = "02";
                                if ((System.String)orows["Ref2"] == "PA")
                                    oTransferStock.ToWarehouse = "03";
                                else if ((System.String)orows["Ref2"] == "B07")
                                    oTransferStock.ToWarehouse = "07";
                                else if ((System.String)orows["Ref2"] == "B99")
                                    oTransferStock.ToWarehouse = "99";
                                else if ((System.String)orows["Ref2"] == "CENABAST")
                                    oTransferStock.ToWarehouse = "06";
                                else if ((System.String)orows["Ref2"] == "VTA")
                                    oTransferStock.ToWarehouse = "08";
                                else if ((System.String)orows["Ref2"] == "LOTE")
                                    oTransferStock.ToWarehouse = "08";
                                else if ((System.String)orows["Ref2"] == "PRILOGIC")
                                    oTransferStock.ToWarehouse = "09";
                                oTransferStock.Lines.ItemCode = (System.String)orows["ItemCode"];
                                oTransferStock.Lines.Quantity = (System.Double)orows["Cantidad"];

                                if (bConLote == true)
                                {
                                    oTransferStock.Lines.BatchNumbers.Quantity = (System.Double)orows["Cantidad"];
                                    oTransferStock.Lines.BatchNumbers.BatchNumber = (System.String)orows["Ref4"];
                                }

                                lRetCode = oTransferStock.Add();
                                if (lRetCode != 0)
                                {
                                    oCompany.GetLastError(out lErrcode, out sErrmsg);
                                    Func.AddLog("Error generar transferencia para " + sv + " - " + sErrmsg );
                                    Func.AddLogSap(oCompany, "Error generar transferencia para " + sv + " - " + sErrmsg, sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                   , (System.String)orows["Razon"], (System.String)orows["Ref1"], (System.String)orows["Ref2"], (System.String)orows["Ref4"], "", (System.String)orows["ItemCode"], "", "", "");
                                }
                                else
                                {
                                    Func.AddLog("transferencia generado para " + sv + " -  tipo 23");
                                    Func.AddLogSap(oCompany, "Transferencia generado para " + sv + "tipo 23", sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                   , (System.String)orows["Razon"], (System.String)orows["Ref1"], (System.String)orows["Ref2"], (System.String)orows["Ref4"], "", (System.String)orows["ItemCode"], "", "", "");
                                }
                            }
                        }
                        else if (((System.String)orows["Code"] == "23") && ((System.String)orows["CodBloqueo"] == "PP") && ((System.String)orows["Razon"] == "CLPN"))
                        {
                            if ((System.Double)orows["Cantidad"] > 0)
                            {
                                bGenerarAjuste = false;
                                if (bConLote == false)
                                {
                                    s = @"select OnHand 'Quantity' from OITW where ItemCode = '{0}' and WhsCode = '{1}' and OnHand > 0";
                                    s = String.Format(s, (System.String)orows["ItemCode"], "03");
                                }
                                else
                                {
                                    s = @"select T1.Quantity 
                                        from OBTN T0 
                                        JOIN OBTQ T1 ON T1.ItemCode = T0.ItemCode 
                                                    and T1.SysNumber = T0.SysNumber 
                                       where T0.ItemCode = '{0}' 
                                         and T1.WhsCode = '{2}' 
                                         and T0.DistNumber = '{1}'
                                         and T1.Quantity > 0";
                                    s = String.Format(s, (System.String)orows["ItemCode"], (System.String)orows["Ref4"], "03");
                                }

                                oRecordSet.DoQuery(s);
                                if (oRecordSet.RecordCount > 0)
                                {
                                    if (((System.Double)oRecordSet.Fields.Item("Quantity").Value) >= (System.Double)orows["Cantidad"])
                                        bGenerarAjuste = true;

                                    if (bGenerarAjuste)
                                    {
                                        oTransferStock = ((SAPbobsCOM.StockTransfer)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer));
                                        oTransferStock.DocDate = DateTime.Now;
                                        oTransferStock.FromWarehouse = "03";
                                        oTransferStock.ToWarehouse = "02";
                                        oTransferStock.Lines.ItemCode = (System.String)orows["ItemCode"];
                                        oTransferStock.Lines.Quantity = (System.Double)orows["Cantidad"];

                                        if (bConLote == true)
                                        {
                                            oTransferStock.Lines.BatchNumbers.Quantity = (System.Double)orows["Cantidad"];
                                            oTransferStock.Lines.BatchNumbers.BatchNumber = (System.String)orows["Ref4"];
                                        }

                                        lRetCode = oTransferStock.Add();
                                        if (lRetCode != 0)
                                        {
                                            oCompany.GetLastError(out lErrcode, out sErrmsg);
                                            Func.AddLog("Error generar transferencia para " + sv + " - " + sErrmsg);
                                            Func.AddLogSap(oCompany, "Error generar transferencia para " + sv + " - " + sErrmsg, sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                           , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", (System.String)orows["Ref4"], "", (System.String)orows["ItemCode"], "", "", "");
                                        }
                                        else
                                        {
                                            Func.AddLog("transferencia generado para " + sv + " -  tipo 23");
                                            Func.AddLogSap(oCompany, "Transferencia generado para " + sv + "tipo 23", sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                           ,(System.String)orows["Razon"], (System.String)orows["Ref1"], "", (System.String)orows["Ref4"], "", (System.String)orows["ItemCode"], "", "", "");
                                        }
                                    }
                                    else
                                    {
                                        oDocuments = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(BoObjectTypes.oInventoryGenEntry);
                                        oDocuments.DocDate = DateTime.Now;
                                        oDocuments.Lines.ItemCode = (System.String)orows["ItemCode"];
                                        oDocuments.Lines.Quantity = (System.Double)orows["Cantidad"];
                                        oDocuments.Lines.WarehouseCode = "02";

                                        if (bConLote == true)
                                        {
                                            oDocuments.Lines.BatchNumbers.Quantity = (System.Double)orows["Cantidad"];
                                            oDocuments.Lines.BatchNumbers.BatchNumber = (System.String)orows["Ref4"];
                                        }

                                        lRetCode = oDocuments.Add();
                                        if (lRetCode != 0)
                                        {
                                            oCompany.GetLastError(out lErrcode, out sErrmsg);
                                            Func.AddLog("Error generar ajuste para " + sv + " - " + sErrmsg);
                                            Func.AddLogSap(oCompany, "Error generar ajuste para " + sv + " - " + sErrmsg, sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                           , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", (System.String)orows["Ref4"], "", (System.String)orows["ItemCode"], "", "", "");
                                        }
                                        else
                                        {
                                            Func.AddLog("Ajuste generado para " + sv + " -  tipo 23");
                                            Func.AddLogSap(oCompany, "Ajuste generado para " + sv + " -  tipo 23", sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                           , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", (System.String)orows["Ref4"], "", (System.String)orows["ItemCode"], "", "", "");
                                        }
                                    }
                                }
                                else
                                {
                                    oDocuments = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(BoObjectTypes.oInventoryGenEntry);
                                    oDocuments.DocDate = DateTime.Now;
                                    oDocuments.Lines.ItemCode = (System.String)orows["ItemCode"];
                                    oDocuments.Lines.Quantity = (System.Double)orows["Cantidad"];
                                    oDocuments.Lines.WarehouseCode = "02";
                                    if (bConLote == true)
                                    {
                                        oDocuments.Lines.BatchNumbers.Quantity = (System.Double)orows["Cantidad"];
                                        oDocuments.Lines.BatchNumbers.BatchNumber = (System.String)orows["Ref4"];
                                    }

                                    lRetCode = oDocuments.Add();
                                    if (lRetCode != 0)
                                    {
                                        oCompany.GetLastError(out lErrcode, out sErrmsg);
                                        Func.AddLog("Error generar ajuste para " + sv + " - " + sErrmsg);
                                        Func.AddLogSap(oCompany, "Error generar ajuste para " + sv + " - " + sErrmsg, sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                       , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", (System.String)orows["Ref4"], "", (System.String)orows["ItemCode"], "", "", "");
                                    }
                                    else
                                    {
                                        Func.AddLog("Ajuste generado para " + sv + " -  tipo 23");
                                        Func.AddLogSap(oCompany, "Ajuste generado para " + sv + " -  tipo 23", sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                       , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", (System.String)orows["Ref4"], "", (System.String)orows["ItemCode"], "", "", "");
                                    }
                                }
                            }
                        }
                        else if (((System.String)orows["Code"] == "25") && ((System.String)orows["CodBloqueo"] != "PA") && ((System.String)orows["Razon"] == "") && (((System.String)orows["Ref1"] == "B07") ||
                            ((System.String)orows["Ref1"] == "B99") || ((System.String)orows["Ref1"] == "CENABAST") 
                            || ((System.String)orows["Ref1"] == "VTA") || ((System.String)orows["Ref1"] == "LOTE")  ||
                            ((System.String)orows["Ref1"] == "PRILOGIC")))
                        {
                            if ((System.Double)orows["Cantidad"] > 0)
                            {

                                oTransferStock = ((SAPbobsCOM.StockTransfer)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer));
                                oTransferStock.DocDate = DateTime.Now;

                                if ((System.String)orows["Ref1"] == "B07")
                                    oTransferStock.FromWarehouse = "07";
                                else if ((System.String)orows["Ref1"] == "B99")
                                    oTransferStock.FromWarehouse = "99";
                                else if ((System.String)orows["Ref1"] == "CENABAST")
                                    oTransferStock.FromWarehouse = "06";
                                else if ((System.String)orows["Ref1"] == "VTA")
                                    oTransferStock.FromWarehouse = "08";
                                else if ((System.String)orows["Ref1"] == "LOTE")
                                    oTransferStock.FromWarehouse = "08";
                                else if ((System.String)orows["Ref1"] == "PRILOGIC")
                                    oTransferStock.FromWarehouse = "09";
                                oTransferStock.ToWarehouse = "02";
                                oTransferStock.Lines.ItemCode = (System.String)orows["ItemCode"];
                                oTransferStock.Lines.Quantity = (System.Double)orows["Cantidad"];

                                if (bConLote == true)
                                {
                                    oTransferStock.Lines.BatchNumbers.Quantity = (System.Double)orows["Cantidad"];
                                    oTransferStock.Lines.BatchNumbers.BatchNumber = (System.String)orows["Ref4"];
                                }

                                lRetCode = oTransferStock.Add();
                                if (lRetCode != 0)
                                {
                                    oCompany.GetLastError(out lErrcode, out sErrmsg);
                                    string descripcion = oCompany.GetLastErrorDescription();
                                    Func.AddLog("Error generar transferencia para " + sv + " - " + sErrmsg + " Des " + descripcion);
                                    Func.AddLogSap(oCompany, "Error generar transferencia para " + sv + " - " + sErrmsg, sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                   , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", (System.String)orows["Ref4"], "", (System.String)orows["ItemCode"], "", "", "");
                                }
                                else
                                {
                                    Func.AddLog("Transferencia generado para " + sv + " -  tipo 25");
                                    Func.AddLogSap(oCompany, "Transferencia generado para " + sv + "tipo 25", sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                   ,(System.String)orows["Razon"], (System.String)orows["Ref1"], "", (System.String)orows["Ref4"], "", (System.String)orows["ItemCode"], "", "", "");
                                }
                            }
                        }
                        else if (((System.String)orows["Code"] == "39") && (((System.String)orows["CodBloqueo"] == "") || ((System.String)orows["CodBloqueo"] == "B07") || ((System.String)orows["CodBloqueo"] == "B99")
                            || ((System.String)orows["CodBloqueo"] == "CENABAST") 
                            || ((System.String)orows["CodBloqueo"] == "VTA") || ((System.String)orows["CodBloqueo"] == "LOTE") || ((System.String)orows["CodBloqueo"] == "PRILOGIC")))
                        {
                            if ((System.Double)orows["Cantidad"] > 0)
                            {
                                bGenerarAjuste = false;
                                if (bConLote == false)
                                {
                                    s = @"select OnHand 'Quantity' from OITW where ItemCode = '{0}' and WhsCode = '{1}' and OnHand > 0";
                                    s = String.Format(s, (System.String)orows["ItemCode"], "03");
                                }
                                else
                                {
                                    s = @"select T1.Quantity 
                                        from OBTN T0 
                                        JOIN OBTQ T1 ON T1.ItemCode = T0.ItemCode 
                                                    and T1.SysNumber = T0.SysNumber 
                                       where T0.ItemCode = '{0}' 
                                         and T1.WhsCode = '{2}' 
                                         and T0.DistNumber = '{1}'
                                         and T1.Quantity > 0";
                                    s = String.Format(s, (System.String)orows["ItemCode"], (System.String)orows["Ref1"], "03");
                                }

                                oRecordSet.DoQuery(s);
                                if (oRecordSet.RecordCount > 0)
                                {
                                    if (((System.Double)oRecordSet.Fields.Item("Quantity").Value) >= (System.Double)orows["Cantidad"])
                                        bGenerarAjuste = true;

                                    if (bGenerarAjuste)
                                    {
                                        oTransferStock = ((SAPbobsCOM.StockTransfer)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer));
                                        oTransferStock.DocDate = DateTime.Now;
                                        oTransferStock.FromWarehouse = "03";
                                        if ((System.String)orows["CodBloqueo"] == "")
                                            oTransferStock.ToWarehouse = "02";
                                        else if ((System.String)orows["CodBloqueo"] == "B07")
                                            oTransferStock.ToWarehouse = "07";
                                        else if ((System.String)orows["CodBloqueo"] == "B99")
                                            oTransferStock.ToWarehouse = "99";
                                        else if ((System.String)orows["CodBloqueo"] == "CENABAST")
                                            oTransferStock.FromWarehouse = "06";
                                        else if ((System.String)orows["CodBloqueo"] == "VTA")
                                            oTransferStock.FromWarehouse = "08";
                                        else if ((System.String)orows["CodBloqueo"] == "LOTE")
                                            oTransferStock.FromWarehouse = "08";
                                        else if ((System.String)orows["CodBloqueo"] == "PRILOGIC")
                                            oTransferStock.FromWarehouse = "09";

                                        oTransferStock.Lines.ItemCode = (System.String)orows["ItemCode"];
                                        oTransferStock.Lines.Quantity = (System.Double)orows["Cantidad"];

                                        if (bConLote == true)
                                        {
                                            oTransferStock.Lines.BatchNumbers.Quantity = (System.Double)orows["Cantidad"];
                                            oTransferStock.Lines.BatchNumbers.BatchNumber = (System.String)orows["Ref1"];
                                        }

                                        lRetCode = oTransferStock.Add();
                                        if (lRetCode != 0)
                                        {
                                            oCompany.GetLastError(out lErrcode, out sErrmsg);
                                            Func.AddLog("Error generar transferencia para " + sv + " - " + sErrmsg);
                                            Func.AddLogSap(oCompany, "Error generar transferencia para " + sv + " - " + sErrmsg, sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                           , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", "", "", (System.String)orows["ItemCode"], "", "", "");
                                        }
                                        else
                                        {
                                            Func.AddLog("Transferencia generado para " + sv + " -  tipo 39");
                                            Func.AddLogSap(oCompany, "Transferencia generado para " + sv + "tipo 39", sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                           , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", "", "", (System.String)orows["ItemCode"], "", "", "");
                                        }
                                    }

                                }
                                else
                                {
                                    oDocuments = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(BoObjectTypes.oInventoryGenEntry);
                                    oDocuments.DocDate = DateTime.Now;
                                    oDocuments.Lines.ItemCode = (System.String)orows["ItemCode"];
                                    oDocuments.Lines.Quantity = (System.Double)orows["Cantidad"];
                                    if ((System.String)orows["CodBloqueo"] == "")
                                        oDocuments.Lines.WarehouseCode = "02";
                                    else if ((System.String)orows["CodBloqueo"] == "B07")
                                        oDocuments.Lines.WarehouseCode = "07";
                                    else if ((System.String)orows["CodBloqueo"] == "B99")
                                        oDocuments.Lines.WarehouseCode = "99";
                                    else if ((System.String)orows["CodBloqueo"] == "CENABAST")
                                        oDocuments.Lines.WarehouseCode = "06";
                                    else if ((System.String)orows["CodBloqueo"] == "VTA")
                                        oDocuments.Lines.WarehouseCode = "08";
                                    else if ((System.String)orows["CodBloqueo"] == "LOTE")
                                        oDocuments.Lines.WarehouseCode = "08";
                                    else if ((System.String)orows["CodBloqueo"] == "PRILOGIC")
                                        oDocuments.Lines.WarehouseCode = "09";

                                    if (bConLote == true)
                                    {
                                        oDocuments.Lines.BatchNumbers.Quantity = (System.Double)orows["Cantidad"];
                                        oDocuments.Lines.BatchNumbers.BatchNumber = (System.String)orows["Ref1"];
                                    }

                                    lRetCode = oDocuments.Add();
                                    if (lRetCode != 0)
                                    {
                                        oCompany.GetLastError(out lErrcode, out sErrmsg);
                                        Func.AddLog("Error generar ajuste para " + sv + " - " + sErrmsg);
                                        Func.AddLogSap(oCompany, "Error generar ajuste para " + sv + " - " + sErrmsg, sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                       , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", "", "", (System.String)orows["ItemCode"], "", "", "");
                                    }
                                    else
                                    {
                                        Func.AddLog("Ajuste generado para " + sv + " -  tipo 39");
                                        Func.AddLogSap(oCompany, "Ajuste generado para " + sv + " -  tipo 39", sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                       , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", "", "", (System.String)orows["ItemCode"], "", "", "");
                                    }

                                }
                            }
                        }
                        else if (((System.String)orows["Code"] == "40") && (((System.String)orows["CodBloqueo"] == "") || ((System.String)orows["CodBloqueo"] == "PA") || ((System.String)orows["CodBloqueo"] == "B07") || ((System.String)orows["CodBloqueo"] == "B99")
                                || ((System.String)orows["CodBloqueo"] == "CENABAST") 
                                || ((System.String)orows["CodBloqueo"] == "VTA") || ((System.String)orows["CodBloqueo"] == "LOTE") || ((System.String)orows["CodBloqueo"] == "PRILOGIC")))
                        {
                            if ((System.Double)orows["Cantidad"] > 0)
                            {

                                oTransferStock = ((SAPbobsCOM.StockTransfer)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer));
                                oTransferStock.DocDate = DateTime.Now;
                                if (((System.String)orows["CodBloqueo"] == "") || ((System.String)orows["CodBloqueo"] == "PA"))
                                    oTransferStock.FromWarehouse = "02";
                                else if ((System.String)orows["CodBloqueo"] == "B07")
                                    oTransferStock.FromWarehouse = "07";
                                else if ((System.String)orows["CodBloqueo"] == "B99")
                                    oTransferStock.FromWarehouse = "99";
                                else if ((System.String)orows["CodBloqueo"] == "CENABAST")
                                    oTransferStock.FromWarehouse = "06";
                                else if ((System.String)orows["CodBloqueo"] == "VTA")
                                    oTransferStock.FromWarehouse = "08";
                                else if ((System.String)orows["CodBloqueo"] == "LOTE")
                                    oTransferStock.FromWarehouse = "08";
                                else if ((System.String)orows["CodBloqueo"] == "PRILOGIC")
                                    oTransferStock.FromWarehouse = "09";

                                oTransferStock.ToWarehouse = "03";
                                oTransferStock.Lines.ItemCode = (System.String)orows["ItemCode"];
                                oTransferStock.Lines.Quantity = (System.Double)orows["Cantidad"];

                                if (bConLote == true)
                                {
                                    oTransferStock.Lines.BatchNumbers.Quantity = (System.Double)orows["Cantidad"];
                                    oTransferStock.Lines.BatchNumbers.BatchNumber = (System.String)orows["Ref1"];
                                }

                                lRetCode = oTransferStock.Add();
                                if (lRetCode != 0)
                                {
                                    oCompany.GetLastError(out lErrcode, out sErrmsg);
                                    Func.AddLog("Error generar transferencia para " + sv + " - " + sErrmsg);
                                    Func.AddLogSap(oCompany, "Error generar transferencia para " + sv + " - " + sErrmsg, sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                   , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", "", "", (System.String)orows["ItemCode"], "", "", "");
                                }
                                else
                                {
                                    Func.AddLog("Transferencia generado para " + sv + " -  tipo 40");
                                    Func.AddLogSap(oCompany, "Transferencia generado para " + sv + "tipo 40", sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                   , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", "", "", (System.String)orows["ItemCode"], "", "", "");
                                }
                            }
                        }
                        if (((System.String)orows["Code"] == "53") && (((System.String)orows["CodBloqueo"] == "") || ((System.String)orows["CodBloqueo"] == "B07") || ((System.String)orows["CodBloqueo"] == "B99")
                           || ((System.String)orows["CodBloqueo"] == "CENABAST") 
                           || ((System.String)orows["CodBloqueo"] == "VTA") || ((System.String)orows["CodBloqueo"] == "LOTE") || ((System.String)orows["CodBloqueo"] == "PRILOGIC")))
                        {
                            if ((System.Double)orows["Cantidad"] > 0)
                            {
                                bGenerarAjuste = false;
                                if (bConLote == false)
                                {
                                    s = @"select OnHand 'Quantity' from OITW where ItemCode = '{0}' and WhsCode = '{1}' and OnHand > 0";
                                    s = String.Format(s, (System.String)orows["ItemCode"], "03");
                                }
                                else
                                {
                                    s = @"select T1.Quantity 
                                        from OBTN T0 
                                        JOIN OBTQ T1 ON T1.ItemCode = T0.ItemCode 
                                                    and T1.SysNumber = T0.SysNumber 
                                       where T0.ItemCode = '{0}' 
                                         and T1.WhsCode = '{2}' 
                                         and T0.DistNumber = '{1}'
                                         and T1.Quantity > 0";
                                    s = String.Format(s, (System.String)orows["ItemCode"], (System.String)orows["Ref1"], "03");
                                }

                                oRecordSet.DoQuery(s);
                                if (oRecordSet.RecordCount > 0)
                                {
                                    if (((System.Double)oRecordSet.Fields.Item("Quantity").Value) >= (System.Double)orows["Cantidad"])
                                        bGenerarAjuste = true;

                                    if (bGenerarAjuste)
                                    {
                                        oTransferStock = ((SAPbobsCOM.StockTransfer)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer));
                                        oTransferStock.DocDate = DateTime.Now;
                                        oTransferStock.FromWarehouse = "03";
                                        if ((System.String)orows["CodBloqueo"] == "")
                                            oTransferStock.ToWarehouse = "02";
                                        else if ((System.String)orows["CodBloqueo"] == "B07")
                                            oTransferStock.ToWarehouse = "07";
                                        else if ((System.String)orows["CodBloqueo"] == "B99")
                                            oTransferStock.ToWarehouse = "99";
                                        else if ((System.String)orows["CodBloqueo"] == "CENABAST")
                                            oTransferStock.ToWarehouse = "06";
                                        else if ((System.String)orows["CodBloqueo"] == "VTA")
                                            oTransferStock.ToWarehouse = "08";
                                        else if ((System.String)orows["CodBloqueo"] == "LOTE")
                                            oTransferStock.ToWarehouse = "08";
                                        else if ((System.String)orows["CodBloqueo"] == "PRILOGIC")
                                            oTransferStock.ToWarehouse = "09";

                                        oTransferStock.Lines.ItemCode = (System.String)orows["ItemCode"];
                                        oTransferStock.Lines.Quantity = (System.Double)orows["Cantidad"];

                                        if (bConLote == true)
                                        {
                                            oTransferStock.Lines.BatchNumbers.Quantity = (System.Double)orows["Cantidad"];
                                            oTransferStock.Lines.BatchNumbers.BatchNumber = (System.String)orows["Ref1"];
                                        }

                                        lRetCode = oTransferStock.Add();
                                        if (lRetCode != 0)
                                        {
                                            oCompany.GetLastError(out lErrcode, out sErrmsg);
                                            Func.AddLog("Error generar transferencia para " + sv + " - " + sErrmsg);
                                            Func.AddLogSap(oCompany, "Error generar transferencia para " + sv + " - " + sErrmsg, sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                           , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", "", "", (System.String)orows["ItemCode"], "", "", "");
                                        }
                                        else
                                        {
                                            Func.AddLog("Transferencia generado para " + sv + " -  tipo 53");
                                            Func.AddLogSap(oCompany, "Transferencia generado para " + sv + "tipo 53", sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                           , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", "", "", (System.String)orows["ItemCode"], "", "", "");
                                        }
                                    }

                                }
                                else
                                {
                                    oDocuments = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(BoObjectTypes.oInventoryGenEntry);
                                    oDocuments.DocDate = DateTime.Now;
                                    oDocuments.Lines.ItemCode = (System.String)orows["ItemCode"];
                                    oDocuments.Lines.Quantity = (System.Double)orows["Cantidad"];
                                    if ((System.String)orows["CodBloqueo"] == "")
                                        oDocuments.Lines.WarehouseCode = "02";
                                    else if ((System.String)orows["CodBloqueo"] == "B07")
                                        oDocuments.Lines.WarehouseCode = "07";
                                    else if ((System.String)orows["CodBloqueo"] == "B99")
                                        oDocuments.Lines.WarehouseCode = "99";
                                    else if ((System.String)orows["CodBloqueo"] == "CENABAST")
                                        oDocuments.Lines.WarehouseCode = "06";
                                    else if ((System.String)orows["CodBloqueo"] == "VTA")
                                        oDocuments.Lines.WarehouseCode = "08";
                                    else if ((System.String)orows["CodBloqueo"] == "LOTE")
                                        oDocuments.Lines.WarehouseCode = "08";
                                    else if ((System.String)orows["CodBloqueo"] == "PRILOGIC")
                                        oDocuments.Lines.WarehouseCode = "09";

                                    if (bConLote == true)
                                    {
                                        oDocuments.Lines.BatchNumbers.Quantity = (System.Double)orows["Cantidad"];
                                        oDocuments.Lines.BatchNumbers.BatchNumber = (System.String)orows["Ref1"];
                                    }

                                    lRetCode = oDocuments.Add();
                                    if (lRetCode != 0)
                                    {
                                        oCompany.GetLastError(out lErrcode, out sErrmsg);
                                        Func.AddLog("Error generar ajuste para " + sv + " - " + sErrmsg);
                                        Func.AddLogSap(oCompany, "Error generar ajuste para " + sv + " - " + sErrmsg, sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                       , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", "", "", (System.String)orows["ItemCode"], "", "", "");
                                    }
                                    else
                                    {
                                        Func.AddLog("Ajuste generado para " + sv + " -  tipo 53");
                                        Func.AddLogSap(oCompany, "Ajuste generado para " + sv + " -  tipo 53", sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                       , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", "", "", (System.String)orows["ItemCode"], "", "", "");
                                    }

                                }
                            }
                            else
                            {
                                oTransferStock = ((SAPbobsCOM.StockTransfer)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer));
                                oTransferStock.DocDate = DateTime.Now;
                                if ((System.String)orows["CodBloqueo"] == "")
                                    oTransferStock.FromWarehouse = "02";
                                else if ((System.String)orows["CodBloqueo"] == "B07")
                                    oTransferStock.FromWarehouse = "07";
                                else if ((System.String)orows["CodBloqueo"] == "B99")
                                    oTransferStock.FromWarehouse = "99";
                                else if ((System.String)orows["CodBloqueo"] == "CENABAST")
                                    oTransferStock.FromWarehouse = "06";
                                else if ((System.String)orows["CodBloqueo"] == "VTA")
                                    oTransferStock.FromWarehouse = "08";
                                else if ((System.String)orows["CodBloqueo"] == "LOTE")
                                    oTransferStock.FromWarehouse = "08";
                                else if ((System.String)orows["CodBloqueo"] == "PRILOGIC")
                                    oTransferStock.FromWarehouse = "09";

                                oTransferStock.ToWarehouse = "03";
                                oTransferStock.Lines.ItemCode = (System.String)orows["ItemCode"];
                                oTransferStock.Lines.Quantity = (System.Double)orows["Cantidad"] * -1;

                                if (bConLote == true)
                                {
                                    oTransferStock.Lines.BatchNumbers.Quantity = (System.Double)orows["Cantidad"] * -1;
                                    oTransferStock.Lines.BatchNumbers.BatchNumber = (System.String)orows["Ref1"];
                                }

                                lRetCode = oTransferStock.Add();
                                if (lRetCode != 0)
                                {
                                    oCompany.GetLastError(out lErrcode, out sErrmsg);
                                    Func.AddLog("Error generar transferencia para " + sv + " - " + sErrmsg);
                                    Func.AddLogSap(oCompany, "Error generar transferencia para " + sv + " - " + sErrmsg, sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                   , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", "", "", (System.String)orows["ItemCode"], "", "", "");
                                }
                                else
                                {
                                    Func.AddLog("transferencia generado para " + sv + " -  tipo 53");
                                    Func.AddLogSap(oCompany, "Transferencia generado para " + sv + "tipo 53", sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                   ,(System.String)orows["Razon"], (System.String)orows["Ref1"], "", "", "", (System.String)orows["ItemCode"], "", "", "");
                                }
                            }
                        }
                        else if (((System.String)orows["Code"] == "56") && ((System.String)orows["CodBloqueo"] == "LOTE"))
                        {
                            s = @"update OBTN set Status = '1' where ItemCode = '{0}' and DistNumber = '{1}'";
                            s = String.Format(s, (System.String)orows["ItemCode"], (System.String)orows["Ref1"]);
                            oRecordSet.DoQuery(s);
                            Func.AddLog("Lote bloqueado para " + sv + " -  tipo 56");
                        }
                        else if (((System.String)orows["Code"] == "56") && ((System.String)orows["CodBloqueo"] == ""))
                        {
                            s = @"update OBTN set Status = '0' where ItemCode = '{0}' and DistNumber = '{1}'";
                            s = String.Format(s, (System.String)orows["ItemCode"], (System.String)orows["Ref1"]);
                            oRecordSet.DoQuery(s);
                            Func.AddLog("Lote liberado  para " + sv + " -  tipo 56");
                        }
                        else if (((System.String)orows["Code"] == "27") && ((System.String)orows["CodBloqueo"] == ""))
                        {
                            oDocuments = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(BoObjectTypes.oOrders);
                            s = ((System.String)orows["Orden"]).Replace("CLI", "");
                            if (s.IndexOf("-") != -1)
                                s = s.Substring(0, s.IndexOf("-"));
                            else
                                s = s.Substring(0, s.Length);

                            if (oDocuments.GetByKey(Convert.ToInt32(s)))
                            {
                                lRetCode = oDocuments.Close();
                                if (lRetCode != 0)
                                {
                                    oCompany.GetLastError(out lErrcode, out sErrmsg);
                                    Func.AddLog("Error cerrar OV para " + sv + " - " + sErrmsg);
                                    Func.AddLogSap(oCompany, "Error cerrar OV para " + sv + " - " + sErrmsg, sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                   , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", "", "", (System.String)orows["ItemCode"], "", "", "");
                                }
                                else
                                {
                                    Func.AddLog("Se cerro OV para " + sv + " -  tipo 27");
                                    Func.AddLogSap(oCompany, "Se cerro OV para " + sv + " -  tipo 27", sv, (System.Double)orows["Cantidad"], (System.String)orows["Code"]
                                                   , (System.String)orows["Razon"], (System.String)orows["Ref1"], "", "", "", (System.String)orows["ItemCode"], "", "", "");
                                }
                            }
                            else
                                Func.AddLog("Error cerrar Documento " + sv + " - No se encuentra documento en SAP" );
                        }
                    }
                    catch (Exception r)
                    {
                        Func.AddLog("Error crear Documento " + sv + " - " + r.Message + ", TRACE " + r.StackTrace);
                    }
                    finally
                    {
                        oDocuments = null;
                        oTransferStock = null;
                    }
                }
                return true;
            }
            catch (Exception w)
            {
                Func.AddLog("Error IHT_Documents - " + w.Message + ", TRACE " + w.StackTrace);
                return false;
            }
        }


        //pensar mejor si recibir los lotes en un arreglo para recorrerlo y cambiar estado y luego deberia volver a llamarse la funciona para dejar estado original
        public void Cambiar_Status_Lote(ref System.Data.DataTable dtlote, Boolean bOriginal)
        {
            TFunctions Func = new TFunctions();
            try
            {
                //Status 0 Liberado
                //Status 1 No Accesible
                //Status 2 Bloqueado
                foreach (System.Data.DataRow orow in dtlote.Rows)
                {
                    s = @"update OBTN set Status = '{2}' where ItemCode = '{0}' and DistNumber = '{1}'";
                    if (bOriginal)
                        s = String.Format(s, ((System.String)orow["ItemCode"]).Trim(), ((System.String)orow["Lote"]).Trim(), ((System.String)orow["Status"]).Trim());
                    else
                        s = String.Format(s, ((System.String)orow["ItemCode"]).Trim(), ((System.String)orow["Lote"]).Trim(), "0");
                    oRecordSet.DoQuery(s);
                }
            }
            catch (Exception te)
            {
                Func.AddLog("Error Cambiar_Status_Lote, " + te.Message + ", TRACE " + te.StackTrace);
            }
        }

        //crea objetos
        public Boolean INS_ResumenInventario(ref System.Data.DataTable dtSorted, String sv)
        {
            Boolean _return = false;
            TFunctions Func = new TFunctions();
            String tablename = "RESWMS";
            SAPbobsCOM.UserTable oUserTable;
            Int32 lRetCode;
            String sErrmsg;
            Int32 lErrcode;
            Int32 iCont;
            try
            {
                //oUserTable = (SAPbobsCOM.UserTable)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables));
                oUserTable = oCompany.UserTables.Item(tablename);
                //primero eliminar registros en la tabla
                s = "select Code from [@{0}]";
                s = String.Format(s, tablename);
                oRecordSet.DoQuery(s);
                while (!oRecordSet.EoF)
                {
                    if (oUserTable.GetByKey(((System.String)oRecordSet.Fields.Item("Code").Value).Trim()))
                        oUserTable.Remove();
                    oRecordSet.MoveNext();
                }

                iCont = 1;
                foreach (System.Data.DataRow orows in dtSorted.Rows)
                {
                    s = "000000000000" + iCont.ToString();
                    s = s.Substring(s.Length - 12, 12);
                    oUserTable = oCompany.UserTables.Item(tablename);
                    oUserTable.Code = s;
                    oUserTable.Name = s;
                    oUserTable.UserFields.Fields.Item("U_Articulo").Value = ((System.String)orows["ItemCode"]);
                    oUserTable.UserFields.Fields.Item("U_Bodega").Value = ((System.String)orows["WhsCode"]);
                    oUserTable.UserFields.Fields.Item("U_Lote").Value = ((System.String)orows["Lote"]);
                    oUserTable.UserFields.Fields.Item("U_UM").Value = ((System.String)orows["UM"]);
                    oUserTable.UserFields.Fields.Item("U_Stock").Value = ((System.Double)orows["Quantity"]).ToString();
                    lRetCode = oUserTable.Add(); //tabla normal

                    if (lRetCode != 0)
                    {
                        oCompany.GetLastError(out lErrcode, out sErrmsg);
                        Func.AddLog("Error guardar registro en " + tablename + " - " + sErrmsg);
                        Func.AddLogSap(oCompany, "Error guardar registro en " + tablename + " - " + sErrmsg, sv, 0, "", "", "", "", "", "", "", "", "", "");
                        break;
                    }
                    iCont++;
                }

                _return = true;
                Func.AddLog("Archivo " + sv + " creado satisfactoriamente en SAP");
                Func.AddLogSap(oCompany, "Archivo " + sv + " creado satisfactoriamente en SAP", sv, 0, "", "", "", "", "", "", "", "", "", "");
               
                return _return;
            }
            catch (Exception wr)
            {
                Func.AddLog("Error INS_ResumenInventario - Archivo " + sv + " - " + wr.Message + ", TRACE " + wr.StackTrace);
                return false;
            }
        }


        public Boolean ConectarBaseSAP()
        {
            XmlDocument xDoc;
            XmlNodeList Configuracion;
            XmlNodeList lista;
            Int32 lRetCode;
            TFunctions Func;
            String sErrMsg;
            String sPath = Path.GetDirectoryName(this.GetType().Assembly.Location);
            Boolean _return = false;

            Func = new TFunctions();
            try
            {
               // Func.AddLog("trata conectar");
                xDoc = new XmlDocument();

                xDoc.Load(sPath + "\\Config.xml");

                Configuracion = xDoc.GetElementsByTagName("Configuracion");
                lista = ((XmlElement)Configuracion[0]).GetElementsByTagName("ServidorSAP");

                foreach (XmlElement nodo in lista)
                {
                    var i = 0;
                    var nServidor = nodo.GetElementsByTagName("Servidor");
                    //string servidor =  nodo.GetElementsByTagName("Servidor");
                    var nLicencia = nodo.GetElementsByTagName("ServLicencia");
                    var nUserSAP = nodo.GetElementsByTagName("UsuarioSAP");
                    var nPassSAP = nodo.GetElementsByTagName("PasswordSAP");
                    var nSQL = nodo.GetElementsByTagName("SQL");
                    var nUserSQL = nodo.GetElementsByTagName("UsuarioSQL");
                    var nPassSQL = nodo.GetElementsByTagName("PasswordSQL");
                    var nBaseSAP = nodo.GetElementsByTagName("BaseSAP");

                    oCompany.Server = (System.String)(nServidor[i].InnerText);
                    oCompany.LicenseServer = (System.String)(nLicencia[i].InnerText);
                    oCompany.DbUserName = (System.String)(nUserSQL[i].InnerText);
                    oCompany.DbPassword = (System.String)(nPassSQL[i].InnerText);

                    if ((System.String)(nSQL[i].InnerText) == "2008")
                        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008;
                    else if ((System.String)(nSQL[i].InnerText) == "2012")
                        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;
                    else if ((System.String)(nSQL[i].InnerText) == "2014")
                        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
                    else if ((System.String)(nSQL[i].InnerText) == "2016")
                        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016;
                    else if ((System.String)(nSQL[i].InnerText) == "2017")
                        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2017;

                    oCompany.UseTrusted = false;
                    oCompany.CompanyDB = (System.String)(nBaseSAP[i].InnerText);
                    oCompany.UserName = (System.String)(nUserSAP[i].InnerText);
                    oCompany.Password = (System.String)(nPassSAP[i].InnerText);
                    /*
                    Func.AddLog(oCompany.Server);
                    Func.AddLog(oCompany.LicenseServer);
                    Func.AddLog(oCompany.DbUserName);
                    //            Func.AddLog(oCompany.DbPassword);
                    Func.AddLog(oCompany.CompanyDB);
                    Func.AddLog(oCompany.UserName);
                    //            Func.AddLog(oCompany.Password);

                    */
                    lRetCode = oCompany.Connect();

                    if (lRetCode != 0)
                    {
                        sErrMsg = oCompany.GetLastErrorDescription();
                        //Func := new TFunciones;
                        Func.AddLog("Error de conexión base SAP, " + sErrMsg);
                        _return = false;
                    }
                    else
                        _return = true;
                }
                return _return;
            }
            catch (Exception w)
            {
                //Func := new TFunciones;
                Func.AddLog("ConectarBase: " + w.Message + " ** Trace: " + w.StackTrace);
                return false;
            }


        }


        private void Download()
        {
            //ssh-rsa 2048 61:82:a1:7e:6b:d5:65:77:ab:1d:bb:92:20:0d:1b:e2
            TFunctions Func = new TFunctions();
            String[] p;
            Session session = new Session();
            Func.AddLog("Inicia sincronizacion");
            
            try
            {
                PathArchivos = Func.PathArchivos();
                // Setup session options
                SessionOptions sessionOptions = new SessionOptions
                {
                    Protocol = Protocol.Sftp,
                    HostName = Func.DatosSFTP("HostName"), //"uat63intf.logfireapps.com"
                    UserName = Func.DatosSFTP("UserName"), //"madegom_63_uat_if"
                    Password = Func.DatosSFTP("Password"), //"wxTTdPN8"
                    SshHostKeyFingerprint = Func.DatosSFTP("SshHostKey") //"ssh-rsa 2048 61:82:a1:7e:6b:d5:65:77:ab:1d:bb:92:20:0d:1b:e2"
                    
                };

                using (session)
                {
                    // Connect
                    session.Open(sessionOptions);
                    Func.AddLog("Conectado");
                    // Upload files
                    TransferOptions transferOptions = new TransferOptions();
                    //transferOptions.TransferMode = TransferMode.Binary;
                    transferOptions.TransferMode = TransferMode.Automatic;

                    string pathDir = Func.DatosSFTP("pathDir");
                    TransferOperationResult transferResult;
                    Func.AddLog("trae archivos");
                    transferResult = session.GetFiles(pathDir + "*.", PathArchivos, false, transferOptions);

                    // Throw on any error
                    transferResult.Check();
                    Func.AddLog("trajo archivos");

                    // Print results
                    foreach (TransferEventArgs transfer in transferResult.Transfers)
                    {
                        p = transfer.FileName.Split('/');
                        //session.MoveFile(transfer.FileName, "/" + p[1] + "/" + p[2] + "/" + p[3] + "/success/" + p[4]);
                        session.RemoveFiles(transfer.FileName);
                        Func.AddLog("Download of " + transfer.FileName + " succeeded");
                    }
                    session.Close();
                }
            }
            catch (Exception ex)
            {
                Func.AddLog("Error sincronizar SFTP - " + ex.Message + ", TRACE " + ex.StackTrace);
                //if (session.Opened)
                //    session.Close();
            }
            finally
            {
                session = null;
                Dispose(true);
                GC.Collect();
            }
            Func.AddLog("Finaliza sincronizacion");
        }



    }
}
