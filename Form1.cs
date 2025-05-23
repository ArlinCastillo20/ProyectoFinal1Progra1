using Newtonsoft.Json;
using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PPT = Microsoft.Office.Interop.PowerPoint;
using CORE = Microsoft.Office.Core; 

namespace Prueba_api
{
    public partial class Form1 : Form
    {

        private const string OpenAIEndpoint = "https://api.openai.com/v1/chat/completions";
        private const string Model = "gpt-4"; // EJEMPLO: "gpt-3.5-turbo" o "gpt-4"
        public Form1()
        {
            InitializeComponent();
        }
        private async void btnEnviar_Click(object sender, EventArgs e)
        {
            string apiKey = txtApiKey.Text;
            string consulta = txtConsulta.Text;

            if (string.IsNullOrEmpty(apiKey))
            {
                MessageBox.Show("Por favor, ingresa tu API Key de OpenAI.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(consulta))
            {
                MessageBox.Show("Por favor, ingresa tu consulta.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            txtResultado.Text = "Cargando...";

            try
            {
                string responseContent = await SendOpenAIChatRequest(apiKey, consulta);
                txtResultado.Text = responseContent;

                string rutaCarpeta = CrearCarpetaResultados();

                GuardarEnBaseDeDatos(consulta, responseContent);

                GenerarArchivoWord(rutaCarpeta, responseContent);

                GenerarPresentacionPowerPoint(rutaCarpeta, responseContent);

            }
            catch (Exception ex)
            {
                txtResultado.Text = $"Error: {ex.Message}";
            }
        }
        private async Task<string> SendOpenAIChatRequest(string apiKey, string prompt)
        {
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);
                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                var requestData = new
                {
                    model = Model,
                    messages = new[]
                    {
                        new { role = "user", content = prompt }
                    }
                };

                string jsonRequest = JsonConvert.SerializeObject(requestData);
                var content = new StringContent(jsonRequest, Encoding.UTF8, "application/json");

                HttpResponseMessage response = await client.PostAsync(OpenAIEndpoint, content);
                response.EnsureSuccessStatusCode();

                string responseJson = await response.Content.ReadAsStringAsync();
                dynamic responseObject = JsonConvert.DeserializeObject(responseJson);

                if (responseObject?.choices?.Count > 0)
                {
                    return responseObject.choices[0].message.content.ToString().Trim();
                }
                else
                {
                    return "No se recibió una respuesta válida de OpenAI.";
                }
            }
        }
        private void GuardarEnBaseDeDatos(string prompt, string resultado)
        {
            
            string connectionString = "Server= DESKTOP-U0HGSAL ;Database= ARLIFINAL1;Integrated Security=True;";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string query = "INSERT INTO Guisel_Asistente (Prompt, Resultado, FechaHora) VALUES (@Prompt, @Resultado, GETDATE())"; 
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Prompt", prompt);
                        command.Parameters.AddWithValue("@Resultado", resultado);

                        int rowsAffected = command.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            Console.WriteLine("Datos guardados en la base de datos."); // Opcional: Para depuración
                        }
                        else
                        {
                            Console.WriteLine("No se pudieron guardar los datos."); // Opcional: Para depuración
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al guardar en la base de datos: {ex.Message}", "Error de Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private string CrearCarpetaResultados()
        {
            string rutaEscritorio = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string rutaCarpeta = System.IO.Path.Combine(rutaEscritorio, "Resultados_Investigacion");
            if (!Directory.Exists(rutaCarpeta))
            {
                Directory.CreateDirectory(rutaCarpeta);
            }
            return rutaCarpeta;
        }
        private void GenerarArchivoWord(string rutaCarpeta, string resultado) //EL WORD FUNCIONA BIEN

        {
            string nombreArchivo = "Resultado_Investigacion.docx";
            string rutaArchivo = System.IO.Path.Combine(rutaCarpeta, nombreArchivo);

            try
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(rutaArchivo, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());

                    DocumentFormat.OpenXml.Wordprocessing.Paragraph para = body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
                    DocumentFormat.OpenXml.Wordprocessing.Run run = para.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run());
                    run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text(resultado));
                }

                MessageBox.Show($"Archivo de Word guardado en: {rutaArchivo}", "Archivo Word Generado", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al generar el archivo de Word: {ex.Message}", "Error Word", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void GenerarPresentacionPowerPoint(string rutaCarpeta, string resultado) //EL POWERPOINT AUN TIENE ERRORES A ARREGLARSE
        {
            string nombreArchivo = "Presentacion_Investigacion.pptx";
            string rutaArchivo = System.IO.Path.Combine(rutaCarpeta, nombreArchivo);

            PPT.Application pptApp = null;
            PPT.Presentation presentation = null;
            PPT.Slide slide = null;

            try
            {
                pptApp = new PPT.Application();
                presentation = pptApp.Presentations.Add(CORE.MsoTriState.msoTrue);
                slide = presentation.Slides.Add(1, PPT.PpSlideLayout.ppLayoutText);
                slide.Shapes[1].TextFrame.TextRange.Text = "Resultado de la Investigación";
                slide.Shapes[2].TextFrame.TextRange.Text = resultado;
                presentation.SaveAs(rutaArchivo, PPT.PpSaveAsFileType.ppSaveAsDefault, CORE.MsoTriState.msoTrue);

                presentation.Close();
                pptApp.Quit();

                MessageBox.Show($"Presentación de PowerPoint guardada en: {rutaArchivo}", "Archivo PowerPoint Generado", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al generar el archivo de PowerPoint: {ex.Message}", "Error PowerPoint", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            { 
                if (slide != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(slide);
                if (presentation != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(presentation);
                if (pptApp != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(pptApp);
            }
        }
    }
}
