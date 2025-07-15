// At the top of ZenifyRibbon.cs
using System;
using System.IO;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Text.Json;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new ZenifyRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace outlook_zenifier
{
    [ComVisible(true)]
    public class ZenifyRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public ZenifyRibbon()
        {
        }

        // Inside the ZenifyRibbon.cs class

        private static readonly HttpClient client = new HttpClient();
        private const string VllmApiEndpoint = "http://localhost:8000/v1/completions"; // CHANGE THIS if your endpoint is different

        // This is the callback specified in the XML's onAction attribute
        public async void OnZenifyButtonClick(Office.IRibbonControl control)
        {
            // 1. Get the active email composer window
            Outlook.Inspector inspector = Globals.ThisAddIn.Application.ActiveInspector();
            if (inspector.CurrentItem is Outlook.MailItem mailItem)
            {
                // 2. Get the Word editor and the user's selection
                Word.Document wordEditor = inspector.WordEditor as Word.Document;
                Word.Selection selection = wordEditor.Application.Selection;
                string selectedText = selection.Text.Trim();

                if (string.IsNullOrEmpty(selectedText))
                {
                    MessageBox.Show("Please select the text you want to 'zenify' first.", "No Text Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                try
                {
                    // 3. Gather all required context
                    string fullEmailBody = wordEditor.Content.Text;
                    string recipients = mailItem.To;
                    string subject = mailItem.Subject;

                    // 4. Construct the prompt for the LLM
                    var prompt = BuildLlmPrompt(fullEmailBody, recipients, subject, selectedText);

                    // 5. Call the LLM and get the zenified text
                    string zenifiedText = await CallVllmApi(prompt);

                    if (string.IsNullOrEmpty(zenifiedText))
                    {
                        MessageBox.Show("The AI model returned an empty response. Please try again.", "API Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // 6. Replace the text in the editor
                    // Strike-through the original selection
                    selection.Font.StrikeThrough = 1; // 1 for true in Word Interop

                    // Collapse the selection to the end and insert the new text
                    selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    selection.TypeText($" [{zenifiedText.Trim()}]");

                    // Important: Un-strikethrough for subsequent typing
                    selection.Font.StrikeThrough = 0;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // Inside the ZenifyRibbon.cs class

        private string BuildLlmPrompt(string fullBody, string recipients, string subject, string selectedText)
        {
            var promptBuilder = new StringBuilder();
            promptBuilder.AppendLine("### Instruction:");
            promptBuilder.AppendLine("You are an AI assistant helping a user write a more diplomatic and professional email. Below is the context of an email correspondence. The user has selected a specific part of their draft to be improved. Your task is to rewrite ONLY the selected text to be more diplomatically acceptable. Respond with ONLY the rewritten text.");
            promptBuilder.AppendLine("\n### Email Context:");
            promptBuilder.AppendLine($"To: {recipients}");
            promptBuilder.AppendLine($"Subject: {subject}");
            promptBuilder.AppendLine("--- Full Email Body ---");
            promptBuilder.AppendLine(fullBody);
            promptBuilder.AppendLine("--- End of Full Email Body ---");
            promptBuilder.AppendLine("\n### Text to Rewrite:");
            promptBuilder.AppendLine(selectedText);
            promptBuilder.AppendLine("\n### Rewritten Text:");

            return promptBuilder.ToString();
        }

        private async Task<string> CallVllmApi(string prompt)
        {
            // The vLLM OpenAI-compatible API expects a specific JSON structure
            var requestData = new
            {
                model = "tiiuae/falcon-7b-instruct", // IMPORTANT: Change to your loaded model name
                prompt = prompt,
                max_tokens = 256, // Adjust as needed
                temperature = 0.7
            };

            var content = new StringContent(JsonSerializer.Serialize(requestData), Encoding.UTF8, "application/json");

            HttpResponseMessage response = await client.PostAsync(VllmApiEndpoint, content);

            if (response.IsSuccessStatusCode)
            {
                string responseBody = await response.Content.ReadAsStringAsync();

                // Parse the JSON to get the text from the first choice
                using (JsonDocument doc = JsonDocument.Parse(responseBody))
                {
                    JsonElement root = doc.RootElement;
                    JsonElement choices = root.GetProperty("choices");
                    if (choices.GetArrayLength() > 0)
                    {
                        return choices[0].GetProperty("text").GetString();
                    }
                }
                return string.Empty;
            }
            else
            {
                // Throw an exception with details from the API
                string errorBody = await response.Content.ReadAsStringAsync();
                throw new HttpRequestException($"API call failed with status code {response.StatusCode}: {errorBody}");
            }
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("outlook_zenifier.ZenifyRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
