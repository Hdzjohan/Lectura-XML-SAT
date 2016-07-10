using LecturaXML.Modelos;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Xml;

namespace LecturaXML
{
    public partial class MainForm : Form
    {
        private bool Base = false;

        private List<Gastos> lista;

        private Microsoft.Office.Interop.Excel.Application docXLS;
        private Workbook libro;
        private Worksheet hojaDeCalculo;

        private XmlDocument docXML;

        public MainForm()
        {
            InitializeComponent();
        }

        // --------------------------------------------------------------------------

        private void btnGenerarGastos_Click(object sender, EventArgs e)
        {
            if (TextBoxReceptor.Text == "")
            {
                MessageBox.Show("El RCF del receptor no puede estar vacio", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();

                openFileDialog.Filter = "xml(*.xml)|*.xml";
                openFileDialog.FilterIndex = 1;
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    lista = new List<Gastos>();
                    foreach (string ruta_archivo in openFileDialog.FileNames)
                    {
                        docXML = new XmlDocument();
                        docXML.Load(@ruta_archivo);
                        XmlNodeList nodoReceptor = docXML.GetElementsByTagName("cfdi:Receptor");
                        if (nodoReceptor.Item(0) != null)
                        {
                            if (TextBoxReceptor.Text == nodoReceptor.Item(0).Attributes["rfc"].Value)
                            {
                                if (Base == false)
                                {
                                    Base = true;
                                    ConstruirXLS(nodoReceptor.Item(0).Attributes["nombre"].Value);
                                }
                                else
                                {
                                    EstablecerDatos();
                                }
                            }
                        }
                    } // Fin foreach

                    ConstruirLista();

                } // Fin openFileDialog.ShowDialog()
            } // Fin else           
        }

        // --------------------------------------------------------------------------

        public void ConstruirXLS(string nombre)
        {
            docXLS = new Microsoft.Office.Interop.Excel.Application();
            docXLS.Visible = true;

            libro = docXLS.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

            hojaDeCalculo = (Worksheet)libro.Worksheets[1];

            Range rango = hojaDeCalculo.Range["A1", "D1"];
            rango.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            rango.Style.Font.Name = "Arial";
            rango.Style.Font.Size = 12;
            rango.Style.Font.Bold = true;

            rango.Merge();
            rango.Value2 = nombre;

            Range rango1 = hojaDeCalculo.Range["A2", "D2"];
            rango1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            rango1.Style.Font.Name = "Arial";
            rango1.Style.Font.Size = 12;
            rango1.Style.Font.Bold = true;

            rango1.Merge();
            rango1.Value2 = "R.F.C " + TextBoxReceptor.Text;

            Range RangoEncabezados1 = hojaDeCalculo.Range["A4"];
            RangoEncabezados1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            RangoEncabezados1.Style.Font.Bold = true;
            RangoEncabezados1.Value2 = "FECHA";
            RangoEncabezados1.EntireColumn.AutoFit();

            Range RangoEncabezados2 = hojaDeCalculo.Range["B4"];
            RangoEncabezados2.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            RangoEncabezados2.Style.Font.Bold = true;
            RangoEncabezados2.Value2 = "CONCEPTO";
            RangoEncabezados2.EntireColumn.AutoFit();

            Range RangoEncabezados3 = hojaDeCalculo.Range["C4"];
            RangoEncabezados3.Style.Font.Bold = true;
            RangoEncabezados3.Value2 = "FOLIO";
            RangoEncabezados3.EntireColumn.AutoFit();

            Range RangoEncabezados4 = hojaDeCalculo.Range["D4"];
            RangoEncabezados4.Style.Font.Bold = true;
            RangoEncabezados4.Value2 = "SUBTOTAL";
            RangoEncabezados4.EntireColumn.AutoFit();
            
            Range RangoEncabezados5 = hojaDeCalculo.Range["E4"];
            RangoEncabezados5.Style.Font.Bold = true;
            RangoEncabezados5.Value2 = "Total impuestos trasladados";
            RangoEncabezados5.EntireColumn.AutoFit();

            Range RangoEncabezados6 = hojaDeCalculo.Range["F4"];
            RangoEncabezados6.Style.Font.Bold = true;
            RangoEncabezados6.Value2 = "IVA";
            RangoEncabezados6.EntireColumn.AutoFit();

            Range RangoEncabezados7 = hojaDeCalculo.Range["G4"];
            RangoEncabezados7.Style.Font.Bold = true;
            RangoEncabezados7.Value2 = "IEPS";
            RangoEncabezados7.EntireColumn.AutoFit();

            Range RangoEncabezados8 = hojaDeCalculo.Range["H4"];
            RangoEncabezados8.Style.Font.Bold = true;
            RangoEncabezados8.Value2 = "TOTAL";
            RangoEncabezados8.EntireColumn.AutoFit();

            Range RangoEncabezados9 = hojaDeCalculo.Range["I4"];
            RangoEncabezados9.Style.Font.Bold = true;
            RangoEncabezados9.Value2 = "DESCRIPCIÓN";
            RangoEncabezados9.EntireColumn.AutoFit();

            Range RangoEncabezados10 = hojaDeCalculo.Range["J4"];
            RangoEncabezados10.Style.Font.Bold = true;
            RangoEncabezados10.Value2 = "CONDICIONES DE PAGO";
            RangoEncabezados10.EntireColumn.AutoFit();

            Range RangoEncabezados11 = hojaDeCalculo.Range["K4"];
            RangoEncabezados11.Style.Font.Bold = true;
            RangoEncabezados11.Value2 = "UUID";
            RangoEncabezados11.EntireColumn.AutoFit();

            EstablecerDatos();
        }

        //--------------------------------------------------------------------------        

        public void EstablecerDatos()
        {
            Gastos gasto = new Gastos();

            XmlNode NodoPrincipal = docXML.DocumentElement;
                        
            gasto.Id = short.Parse(NodoPrincipal.Attributes["fecha"].Value.Substring(8, 2));

            gasto.Fecha = NodoPrincipal.Attributes["fecha"].Value.Substring(0, 10);

            XmlNodeList nodoEmisor = docXML.GetElementsByTagName("cfdi:Emisor");
            if (nodoEmisor.Item(0).Attributes["nombre"] != null)
            {
                gasto.Concepto = nodoEmisor.Item(0).Attributes["nombre"].Value;
            }                          

            string serie = "sin serie", folio = "sin folio";
            if (NodoPrincipal.Attributes["serie"] != null)
            {
                serie = NodoPrincipal.Attributes["serie"].Value;
            }
            if (NodoPrincipal.Attributes["folio"] != null)
            {
                folio = NodoPrincipal.Attributes["folio"].Value;
            }            
            gasto.Folio = serie + " - " + folio;                    

            gasto.Subtotal = NodoPrincipal.Attributes["subTotal"].Value;

            XmlNodeList nodoImpuestos = docXML.GetElementsByTagName("cfdi:Impuestos");
            if (nodoImpuestos.Item(0).Attributes["totalImpuestosTrasladados"] != null)
            {
                gasto.TotalImpuestosTrasladados = nodoImpuestos.Item(0).Attributes["totalImpuestosTrasladados"].Value;
            }            

            XmlNodeList nodoTranslados = docXML.GetElementsByTagName("cfdi:Traslado");            
            if (nodoTranslados.Count != 0)
            {
                if (nodoTranslados.Item(0) != null)
                {
                    gasto.IVA = nodoTranslados.Item(0).Attributes["importe"].Value;
                }                
                if (nodoTranslados.Item(1) != null)
                {
                    gasto.IEPS = nodoTranslados.Item(1).Attributes["importe"].Value;
                }
            }

            gasto.Total = NodoPrincipal.Attributes["total"].Value;

            XmlNodeList nodosConceptos = docXML.GetElementsByTagName("cfdi:Concepto");            
            for (int i = 0; i < nodosConceptos.Count; i++)
            {
                gasto.Descripción += nodosConceptos.Item(i).Attributes["descripcion"].Value + ", ";
            }
            
            if (NodoPrincipal.Attributes["condicionesDePago"] != null)
            {
                gasto.Condiciones = NodoPrincipal.Attributes["condicionesDePago"].Value;
            }

            XmlNodeList nodosUUID = docXML.GetElementsByTagName("tfd:TimbreFiscalDigital");
            gasto.UUID = nodosUUID.Item(0).Attributes["UUID"].Value;

            lista.Add(gasto);
        }

        //--------------------------------------------------------------------------        

        public void ConstruirLista()
        {
            short fila = 5;
            lista.Sort();

            foreach (Gastos gasto in lista)
            {
                Range Rango = hojaDeCalculo.Range["A" + fila.ToString()];
                Rango.Style.Font.Bold = false;
                Rango.Style.Font.Name = "Calibri";
                Rango.Style.Font.Size = 11;
                Rango.Value2 = gasto.Fecha;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["B" + fila.ToString()];
                Rango.Value2 = gasto.Concepto;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["C" + fila.ToString()];
                Rango.Value2 = gasto.Folio;                
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["D" + fila.ToString()];
                Rango.Value2 = gasto.Subtotal;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["E" + fila.ToString()];
                Rango.Value2 = gasto.TotalImpuestosTrasladados;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["F" + fila.ToString()];
                Rango.Value2 = gasto.IVA;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["G" + fila.ToString()];
                Rango.Value2 = gasto.IEPS;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["H" + fila.ToString()];
                Rango.Value2 = gasto.Total;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["I" + fila.ToString()];
                Rango.Value2 = gasto.Descripción;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["J" + fila.ToString()];
                Rango.Value2 = gasto.Condiciones;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["K" + fila.ToString()];
                Rango.Value2 = gasto.UUID;
                Rango.EntireColumn.AutoFit();

                fila++;
            }
        }
    }
}
