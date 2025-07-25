﻿// At the top of ZenifyRibbon.cs
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
using System.Collections.Generic;
using System.Text.Json.Serialization;

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


        private static readonly HttpClient client = new HttpClient();
        private const string ChatEndpoint = "chat/completions"; // The endpoint for Chat Completions

        static ZenifyRibbon()
        {
            // WARNING: The following handler bypasses SSL certificate validation.
            // This is equivalent to `verify=False` in Python's requests/httpx.
            // ONLY use this if you trust the endpoint (e.g., it's on a private network).
            var handler = new HttpClientHandler
            {
                ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => true
            };

            client = new HttpClient(handler);
            client.BaseAddress = new Uri(AppConstants.VllmApiBase);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", AppConstants.VllmApiKey);
        }
        public ZenifyRibbon()
        {
        }

        // Inside the ZenifyRibbon.cs class

        public object getZenifyIcon(Office.IRibbonControl control)
        {
            // The name 'zenify_logo' must match the name of the image
            // in the project's Resources (without the file extension).
            return new System.Drawing.Bitmap(Properties.Resources.zenify_logo);
        }


        // This is the callback specified in the XML's onAction attribute
        public async void OnZenifyButtonClick(Office.IRibbonControl control)
        {
            Outlook.MailItem mailItem = null;
            Word.Document wordEditor = null;

            // Determine the context: Is it a popped-out Inspector or an inline Explorer?
            Outlook.Inspector inspector = Globals.ThisAddIn.Application.ActiveInspector();

            if (inspector != null && inspector.CurrentItem is Outlook.MailItem)
            {
                // CONTEXT 1: We are in a popped-out window (New, Reply, or Forward).
                mailItem = inspector.CurrentItem as Outlook.MailItem;
                wordEditor = inspector.WordEditor as Word.Document;
            }
            else
            {
                // CONTEXT 2: No active inspector. We are likely in an inline response in the main Explorer.
                Outlook.Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer != null && explorer.ActiveInlineResponse != null)
                {
                    // The ActiveInlineResponse is the MailItem being composed.
                    mailItem = explorer.ActiveInlineResponse as Outlook.MailItem;
                    if (mailItem != null)
                    {
                        // Get the Word editor from the MailItem's inspector property, which exists even if not visible.
                        wordEditor = mailItem.GetInspector.WordEditor as Word.Document;
                    }
                }
            }

            // --- The rest of the logic is the same, but it's now wrapped in a null check ---
            if (mailItem != null && wordEditor != null)
            {
                Word.Selection selection = wordEditor.Application.Selection;
                string selectedText = selection.Text.Trim();

                if (string.IsNullOrEmpty(selectedText))
                {
                    MessageBox.Show("Please select the text you want to 'zenify' first.", "No Text Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                try
                {
                    string fullEmailBody = wordEditor.Content.Text;
                    string recipients = mailItem.To;
                    string subject = mailItem.Subject;

                    var userPrompt = BuildLlmPrompt(fullEmailBody, recipients, subject, selectedText);
                    string zenifiedText = await CallVllmChatApi(userPrompt);

                    if (string.IsNullOrEmpty(zenifiedText))
                    {
                        MessageBox.Show("The AI model returned an empty response. Please try again.", "API Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    selection.Font.StrikeThrough = 1; // Strikethrough original
                    selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    selection.TypeText($" [{zenifiedText.Trim()}]");
                    selection.Font.StrikeThrough = 0; // Un-strikethrough for subsequent typing
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred while contacting the AI service:\n\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Could not find an active email editor. Please click inside the message body and try again.", "Editor Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

        // The core API call, now updated for the Chat Completions API
        private async Task<string> CallVllmChatApi(string userPrompt)
        {
            // Create the request payload matching the /chat/completions format
            var requestData = new ChatRequest
            {
                Model = AppConstants.ModelName,
                Messages = new List<ChatMessage>
                {
                    new ChatMessage { Role = "user", Content = userPrompt }
                },
                MaxTokens = 256,
                Temperature = 0.7
            };

            var options = new JsonSerializerOptions { DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull };
            var content = new StringContent(JsonSerializer.Serialize(requestData, options), Encoding.UTF8, "application/json");

            // Post to the correct relative endpoint
            HttpResponseMessage response = await client.PostAsync(ChatEndpoint, content);

            if (response.IsSuccessStatusCode)
            {
                string responseBody = await response.Content.ReadAsStringAsync();
                // Parse the new response structure
                using (JsonDocument doc = JsonDocument.Parse(responseBody))
                {
                    JsonElement root = doc.RootElement;
                    JsonElement choices = root.GetProperty("choices");
                    if (choices.GetArrayLength() > 0)
                    {
                        // The path is now choices -> message -> content
                        return choices[0].GetProperty("message").GetProperty("content").GetString();
                    }
                }
                return string.Empty;
            }
            else
            {
                string responseBody = await response.Content.ReadAsStringAsync();
                // Throw an exception with details from the API for better debugging
                throw new HttpRequestException($"API call failed with status code {response.StatusCode}: {responseBody}");
            }
        }

        // Helper classes to model the JSON payload for the Chat API
        public class ChatRequest
        {
            [JsonPropertyName("model")]
            public string Model { get; set; }

            [JsonPropertyName("messages")]
            public List<ChatMessage> Messages { get; set; }

            [JsonPropertyName("max_tokens")]
            public int? MaxTokens { get; set; }

            [JsonPropertyName("temperature")]
            public double? Temperature { get; set; }
        }

        public class ChatMessage
        {
            [JsonPropertyName("role")]
            public string Role { get; set; }

            [JsonPropertyName("content")]
            public string Content { get; set; }
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
