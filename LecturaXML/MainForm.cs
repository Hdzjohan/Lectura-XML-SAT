using LecturaXML.Modelos;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Xml;

namespace LecturaXML
{
    public partial class MainForm : Form
    {
        private Microsoft.Office.Interop.Excel.Application docXLS;
        private Workbook libro;
        private Worksheet hojaDeCalculo;

        private XmlDocument docXML;        
        private short contadorArchivos = 0;

        private string Nombre = "";

        private List<Proveedor> listaProveedor = null;
        private List<Cliente> listaCliente = null;

        public MainForm()
        {
            InitializeComponent();
        }

        // --------------------------------------------------------------------------

        private void btnGenerarExcel_Click(object sender, EventArgs e)
        {          
            if (TextBoxReceptor.Text == "" || comboBox.SelectedItem == null)
            {
                if (TextBoxReceptor.Text == "")
                {
                    MessageBox.Show(this, "El RCF no puede estar vacio", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                if(comboBox.SelectedItem == null)
                {
                    MessageBox.Show(this, "Seleccione una opción de la lista", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }            
            }            
            else
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();

                openFileDialog.Filter = "xml(*.xml)|*.xml";
                openFileDialog.FilterIndex = 1;
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {                    
                    if(comboBox.SelectedItem.ToString() == "Cliente")
                    {
                        EstablececerDatosCliente(openFileDialog);
                    }
                    else
                    {
                        EstablececerDatosProveedor(openFileDialog);
                    }
                    if (contadorArchivos < openFileDialog.FileNames.Length)
                    {
                        if (listaProveedor !=  null)
                        {
                            GenerarExcelProveedor();                          
                        }
                        else
                        {
                            GenerarExcelCliente();
                        }
                        
                        if (contadorArchivos > 0)
                        {
                            MessageBox.Show(this, contadorArchivos.ToString() + " de los archivos no coincide con RFC", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else
                    {
                        MessageBox.Show(this, "Ninguno de los archivos coincide con el RFC", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                } // Fin openFileDialog.ShowDialog()
            } // Fin else           
        }

        // --------------------------------------------------------------------------

        public void EstablececerDatosCliente(OpenFileDialog openFileDialog)
        {
            listaCliente = new List<Cliente>();
            foreach (string ruta_archivo in openFileDialog.FileNames)
            {
                docXML = new XmlDocument();
                docXML.Load(@ruta_archivo);

                XmlNodeList nodoReceptor = docXML.GetElementsByTagName("cfdi:Emisor");

                if (nodoReceptor.Item(0) != null)
                {
                    if (TextBoxReceptor.Text == nodoReceptor.Item(0).Attributes["rfc"].Value)
                    {
                        Cliente cliente = new Cliente();

                        Nombre = (nodoReceptor.Item(0).Attributes["nombre"] != null) ? nodoReceptor.Item(0).Attributes["nombre"].Value : "Sin nombre";

                        XmlNode NodoPrincipal = docXML.DocumentElement;

                        cliente.Id = short.Parse(NodoPrincipal.Attributes["fecha"].Value.Substring(8, 2));

                        cliente.TipoComprobante = NodoPrincipal.Attributes["tipoDeComprobante"].Value;

                        cliente.Fecha = NodoPrincipal.Attributes["fecha"].Value.Substring(0, 10);

                        cliente.Serie = (NodoPrincipal.Attributes["serie"] != null) ? NodoPrincipal.Attributes["serie"].Value : "Sin serie";

                        cliente.Folio = (NodoPrincipal.Attributes["folio"] != null) ? NodoPrincipal.Attributes["folio"].Value : "Sin folio";

                        cliente.Total = (cliente.TipoComprobante == "egreso") ? "-" + NodoPrincipal.Attributes["total"].Value : NodoPrincipal.Attributes["total"].Value;

                        cliente.Condiciones = (NodoPrincipal.Attributes["condicionesDePago"] != null) ? NodoPrincipal.Attributes["condicionesDePago"].Value : "Sin condiciones";

                        cliente.Subtotal = (cliente.TipoComprobante == "egreso") ? "-" + NodoPrincipal.Attributes["subTotal"].Value : NodoPrincipal.Attributes["subTotal"].Value;

                        XmlNodeList nodoEmisor = docXML.GetElementsByTagName("cfdi:Receptor");

                        cliente.Receptor = (nodoEmisor.Item(0).Attributes["nombre"] != null) ? nodoEmisor.Item(0).Attributes["nombre"].Value : "Sin receptor";
                        cliente.RFC = (nodoEmisor.Item(0).Attributes["rfc"] != null) ? nodoEmisor.Item(0).Attributes["rfc"].Value : "";

                        XmlNodeList nodoImpuestos = docXML.GetElementsByTagName("cfdi:Impuestos");

                        if (nodoImpuestos.Item(0).Attributes["totalImpuestosTrasladados"] != null)
                        {
                            cliente.TotalImpuestosTrasladados = (cliente.TipoComprobante == "egreso") ? "-" + nodoImpuestos.Item(0).Attributes["totalImpuestosTrasladados"].Value : nodoImpuestos.Item(0).Attributes["totalImpuestosTrasladados"].Value;
                        }

                        XmlNodeList nodoTranslados = docXML.GetElementsByTagName("cfdi:Traslado");

                        if (nodoTranslados.Count != 0)
                        {
                            if (nodoTranslados.Item(0) != null)
                            {
                                cliente.IVA = (cliente.TipoComprobante == "egreso") ? "-" + nodoTranslados.Item(0).Attributes["importe"].Value : nodoTranslados.Item(0).Attributes["importe"].Value;
                            }
                            if (nodoTranslados.Item(1) != null)
                            {
                                cliente.IEPS = (cliente.TipoComprobante == "egreso") ? "-" + nodoTranslados.Item(1).Attributes["importe"].Value : nodoTranslados.Item(1).Attributes["importe"].Value;
                            }
                        }

                        XmlNodeList nodosConceptos = docXML.GetElementsByTagName("cfdi:Concepto");

                        for (int i = 0; i < nodosConceptos.Count; i++)
                        {
                            if (nodosConceptos.Count == 1)
                            {
                                cliente.Descripción += nodosConceptos.Item(i).Attributes["descripcion"].Value;
                            }
                            else
                            {
                                if (i == (nodosConceptos.Count - 1))
                                {
                                    cliente.Descripción += nodosConceptos.Item(i).Attributes["descripcion"].Value;
                                }
                                else
                                {
                                    cliente.Descripción += nodosConceptos.Item(i).Attributes["descripcion"].Value + ", ";
                                }
                            }
                        }

                        XmlNodeList nodosUUID = docXML.GetElementsByTagName("tfd:TimbreFiscalDigital");

                        cliente.UUID = nodosUUID.Item(0).Attributes["UUID"].Value;

                        listaCliente.Add(cliente);
                    }
                    else
                    {
                        contadorArchivos++;
                    }
                }
            }
        }

        // --------------------------------------------------------------------------

        public void EstablececerDatosProveedor(OpenFileDialog openFileDialog)
        {
            listaProveedor = new List<Proveedor>();
            foreach (string ruta_archivo in openFileDialog.FileNames)
            {
                docXML = new XmlDocument();
                docXML.Load(@ruta_archivo);

                XmlNodeList nodoReceptor = docXML.GetElementsByTagName("cfdi:Receptor");

                if (nodoReceptor.Item(0) != null)
                {
                    if (TextBoxReceptor.Text == nodoReceptor.Item(0).Attributes["rfc"].Value)
                    {
                        Proveedor proveedor = new Proveedor();

                        Nombre = (nodoReceptor.Item(0).Attributes["nombre"] != null) ? nodoReceptor.Item(0).Attributes["nombre"].Value : "Sin nombre";

                        XmlNode NodoPrincipal = docXML.DocumentElement;

                        proveedor.Id = short.Parse(NodoPrincipal.Attributes["fecha"].Value.Substring(8, 2));

                        proveedor.TipoComprobante = NodoPrincipal.Attributes["tipoDeComprobante"].Value;

                        proveedor.Fecha = NodoPrincipal.Attributes["fecha"].Value.Substring(0, 10);

                        proveedor.Serie = (NodoPrincipal.Attributes["serie"] != null) ? NodoPrincipal.Attributes["serie"].Value : "Sin serie";

                        proveedor.Folio = (NodoPrincipal.Attributes["folio"] != null) ? NodoPrincipal.Attributes["folio"].Value : "Sin folio";

                        proveedor.Total = (proveedor.TipoComprobante == "egreso") ? "-" + NodoPrincipal.Attributes["total"].Value : NodoPrincipal.Attributes["total"].Value;

                        proveedor.Condiciones = (NodoPrincipal.Attributes["condicionesDePago"] != null) ? NodoPrincipal.Attributes["condicionesDePago"].Value : "Sin condiciones";

                        proveedor.Subtotal = (proveedor.TipoComprobante == "egreso") ? "-" + NodoPrincipal.Attributes["subTotal"].Value : NodoPrincipal.Attributes["subTotal"].Value;

                        XmlNodeList nodoEmisor = docXML.GetElementsByTagName("cfdi:Emisor");

                        proveedor.Concepto = (nodoEmisor.Item(0).Attributes["nombre"] != null) ? nodoEmisor.Item(0).Attributes["nombre"].Value : "Sin concepto";

                        XmlNodeList nodoImpuestos = docXML.GetElementsByTagName("cfdi:Impuestos");

                        if (nodoImpuestos.Item(0).Attributes["totalImpuestosTrasladados"] != null)
                        {
                            proveedor.TotalImpuestosTrasladados = (proveedor.TipoComprobante == "egreso") ? "-" + nodoImpuestos.Item(0).Attributes["totalImpuestosTrasladados"].Value : nodoImpuestos.Item(0).Attributes["totalImpuestosTrasladados"].Value;
                        }

                        XmlNodeList nodoTranslados = docXML.GetElementsByTagName("cfdi:Traslado");

                        if (nodoTranslados.Count != 0)
                        {
                            if (nodoTranslados.Item(0) != null)
                            {
                                proveedor.IVA = (proveedor.TipoComprobante == "egreso") ? "-" + nodoTranslados.Item(0).Attributes["importe"].Value : nodoTranslados.Item(0).Attributes["importe"].Value;
                            }
                            if (nodoTranslados.Item(1) != null)
                            {
                                proveedor.IEPS = (proveedor.TipoComprobante == "egreso") ? "-" + nodoTranslados.Item(1).Attributes["importe"].Value : nodoTranslados.Item(1).Attributes["importe"].Value;
                            }
                        }                        

                        XmlNodeList nodosConceptos = docXML.GetElementsByTagName("cfdi:Concepto");

                        for (int i = 0; i < nodosConceptos.Count; i++)
                        {
                            if (nodosConceptos.Count == 1)
                            {
                                proveedor.Descripción += nodosConceptos.Item(i).Attributes["descripcion"].Value;
                            }
                            else
                            {
                                if (i == (nodosConceptos.Count - 1))
                                {
                                    proveedor.Descripción += nodosConceptos.Item(i).Attributes["descripcion"].Value;
                                }
                                else
                                {
                                    proveedor.Descripción += nodosConceptos.Item(i).Attributes["descripcion"].Value + ", ";
                                }
                            }
                        }                        

                        XmlNodeList nodosUUID = docXML.GetElementsByTagName("tfd:TimbreFiscalDigital");

                        proveedor.UUID = nodosUUID.Item(0).Attributes["UUID"].Value;

                        listaProveedor.Add(proveedor);
                    }
                    else
                    {
                        contadorArchivos++;
                    }
                }
            } // fin foreach
        }

        //--------------------------------------------------------------------------        

        public void GenerarExcelCliente()
        {
            docXLS = new Microsoft.Office.Interop.Excel.Application();
            docXLS.Visible = true;

            libro = docXLS.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            hojaDeCalculo = (Worksheet)libro.Worksheets[1];

            Range Rango = hojaDeCalculo.Range["A1", "D1"];
            Rango.Font.Bold = true;
            Rango.Font.Size = 20;
            Rango.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            Rango.Merge();
            Rango.Value2 = Nombre;

            Rango = hojaDeCalculo.Range["A2", "D2"];
            Rango.Font.Bold = true;
            Rango.Font.Size = 18;
            Rango.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            Rango.Merge();
            Rango.Value2 = "R.F.C " + TextBoxReceptor.Text;

            Rango = hojaDeCalculo.Range["A4"];
            Rango.Font.Bold = true;
            Rango.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            Rango.Value2 = "FECHA";
            Rango.EntireColumn.AutoFit();

            Rango = hojaDeCalculo.Range["B4"];
            Rango.Font.Bold = true;
            Rango.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            Rango.Value2 = "Receptor";
            Rango.EntireColumn.AutoFit();

            Rango = hojaDeCalculo.Range["C4"];
            Rango.Font.Bold = true;
            Rango.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            Rango.Value2 = "RFC";
            Rango.EntireColumn.AutoFit();

            Rango = hojaDeCalculo.Range["D4"];
            Rango.Font.Bold = true;
            Rango.Value2 = "FOLIO";
            Rango.EntireColumn.AutoFit();

            Rango = hojaDeCalculo.Range["E4"];
            Rango.Font.Bold = true;
            Rango.Value2 = "SUBTOTAL";
            Rango.EntireColumn.AutoFit();

            Rango = hojaDeCalculo.Range["F4"];
            Rango.Font.Bold = true;
            Rango.Value2 = "Total impuestos trasladados";
            Rango.EntireColumn.AutoFit();

            Rango = hojaDeCalculo.Range["G4"];
            Rango.Font.Bold = true;
            Rango.Value2 = "IVA";
            Rango.EntireColumn.AutoFit();

            Rango = hojaDeCalculo.Range["H4"];
            Rango.Font.Bold = true;
            Rango.Value2 = "IEPS";
            Rango.EntireColumn.AutoFit();

            Rango = hojaDeCalculo.Range["I4"];
            Rango.Font.Bold = true;
            Rango.Value2 = "TOTAL";
            Rango.EntireColumn.AutoFit();

            Rango = hojaDeCalculo.Range["J4"];
            Rango.Font.Bold = true;
            Rango.Value2 = "DESCRIPCIÓN";
            Rango.EntireColumn.AutoFit();

            Rango = hojaDeCalculo.Range["K4"];
            Rango.Font.Bold = true;
            Rango.Value2 = "CONDICIONES DE PAGO";
            Rango.EntireColumn.AutoFit();

            Rango = hojaDeCalculo.Range["L4"];
            Rango.Font.Bold = true;
            Rango.Value2 = "UUID";
            Rango.EntireColumn.AutoFit();

            short fila = 5;
            Color color;

            listaCliente.Sort();

            foreach (Cliente cliente in listaCliente)
            {
                color = (cliente.TipoComprobante == "egreso") ? Color.Red : Color.Black;                

                Rango = hojaDeCalculo.Range["A" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = cliente.Fecha;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["B" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = cliente.Receptor;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["C" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = cliente.RFC;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["D" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = cliente.Serie + " - " + cliente.Folio;                
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["E" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = cliente.Subtotal;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["F" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = cliente.TotalImpuestosTrasladados;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["G" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = cliente.IVA;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["H" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = cliente.IEPS;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["I" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = cliente.Total;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["J" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = cliente.Descripción;
                Rango.EntireColumn.ColumnWidth = 40;

                Rango = hojaDeCalculo.Range["K" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = cliente.Condiciones;
                                
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["L" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = cliente.UUID;
                Rango.EntireColumn.AutoFit();

                fila++;
            }            
        }

        //--------------------------------------------------------------------------        

        public void GenerarExcelProveedor()
        {
            docXLS = new Microsoft.Office.Interop.Excel.Application();
            docXLS.Visible = true;

            libro = docXLS.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            hojaDeCalculo = (Worksheet)libro.Worksheets[1];

            Range Rango = hojaDeCalculo.Range["A1", "D1"];
            Rango.Font.Bold = true;
            Rango.Font.Size = 20;
            Rango.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            Rango.Merge();
            Rango.Value2 = Nombre;

            Rango = hojaDeCalculo.Range["A2", "D2"];
            Rango.Font.Bold = true;
            Rango.Font.Size = 18;
            Rango.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            Rango.Merge();
            Rango.Value2 = "R.F.C " + TextBoxReceptor.Text;

            Rango = hojaDeCalculo.Range["A4"];
            Rango.Font.Bold = true;
            Rango.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            Rango.Value2 = "FECHA";
            Rango.EntireColumn.AutoFit();

            Rango = hojaDeCalculo.Range["B4"];
            Rango.Font.Bold = true;
            Rango.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            Rango.Value2 = "CONCEPTO";
            Rango.EntireColumn.AutoFit();

            Rango = hojaDeCalculo.Range["C4"];
            Rango.Font.Bold = true;
            Rango.Value2 = "FOLIO";
            Rango.EntireColumn.AutoFit();

            Rango = hojaDeCalculo.Range["D4"];
            Rango.Font.Bold = true;
            Rango.Value2 = "SUBTOTAL";
            Rango.EntireColumn.AutoFit();

            Rango = hojaDeCalculo.Range["E4"];
            Rango.Font.Bold = true;
            Rango.Value2 = "Total impuestos trasladados";
            Rango.EntireColumn.AutoFit();

            Rango = hojaDeCalculo.Range["F4"];
            Rango.Font.Bold = true;
            Rango.Value2 = "IVA";
            Rango.EntireColumn.AutoFit();

            Rango = hojaDeCalculo.Range["G4"];
            Rango.Font.Bold = true;
            Rango.Value2 = "IEPS";
            Rango.EntireColumn.AutoFit();

            Rango = hojaDeCalculo.Range["H4"];
            Rango.Font.Bold = true;
            Rango.Value2 = "TOTAL";
            Rango.EntireColumn.AutoFit();

            Rango = hojaDeCalculo.Range["I4"];
            Rango.Font.Bold = true;
            Rango.Value2 = "DESCRIPCIÓN";
            Rango.EntireColumn.AutoFit();

            Rango = hojaDeCalculo.Range["J4"];
            Rango.Font.Bold = true;
            Rango.Value2 = "CONDICIONES DE PAGO";
            Rango.EntireColumn.AutoFit();

            Rango = hojaDeCalculo.Range["K4"];
            Rango.Font.Bold = true;
            Rango.Value2 = "UUID";
            Rango.EntireColumn.AutoFit();

            short fila = 5;
            Color color;

            listaProveedor.Sort();

            foreach (Proveedor proveedor in listaProveedor)
            {
                color = (proveedor.TipoComprobante == "egreso") ? Color.Red : Color.Black;

                Rango = hojaDeCalculo.Range["A" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = proveedor.Fecha;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["B" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = proveedor.Concepto;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["C" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = proveedor.Serie + " - " + proveedor.Folio;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["D" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = proveedor.Subtotal;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["E" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = proveedor.TotalImpuestosTrasladados;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["F" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = proveedor.IVA;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["G" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = proveedor.IEPS;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["H" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = proveedor.Total;
                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["I" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = proveedor.Descripción;
                Rango.EntireColumn.ColumnWidth = 40;

                Rango = hojaDeCalculo.Range["J" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = proveedor.Condiciones;

                Rango.EntireColumn.AutoFit();

                Rango = hojaDeCalculo.Range["K" + fila.ToString()];
                Rango.Font.Color = color;
                Rango.Value2 = proveedor.UUID;
                Rango.EntireColumn.AutoFit();

                fila++;
            }
        }
        //--------------------------------------------------------------------------        
    }
}
