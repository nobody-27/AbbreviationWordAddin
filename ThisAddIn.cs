using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace AbbreviationWordAddin
{
    public partial class ThisAddIn
    {
        public bool abbreviationEnabled = false; // Enable abbreviation replacement by default
        private KeyboardHook keyboardHook;
        public Dictionary<string, string> abbreviations;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WindowSelectionChange += Application_WindowSelectionChange;

            AbbreviationManager.LoadAbbreviations();

            var application = Globals.ThisAddIn.Application;
            AutoCorrect autoCorrect = application.AutoCorrect;
            if (autoCorrect.Equals(true)) {
                abbreviationEnabled = false;
            }
            else
            {
                abbreviationEnabled = true;
            }

            // Add each abbreviation to AutoCorrect
            foreach (var abbreviation in AbbreviationManager.GetAllPhrases())
            {
                // Get the abbreviation and its full form
                string fullForm = AbbreviationManager.GetAbbreviation(abbreviation);

                // Add the abbreviation to AutoCorrect
                this.Application.AutoCorrect.ReplaceText = true; // Ensure AutoCorrect replacement is enabled
                this.Application.AutoCorrect.Entries.Add(abbreviation, fullForm); // Add abbreviation and replacement
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        /// <summary>
        /// Automatically replaces text when the user types if abbreviations are enabled.
        /// </summary>
        private void Application_WindowSelectionChange(Word.Selection sel)
        {
            if (abbreviationEnabled && sel != null && sel.Text.Length > 3)  // Avoid replacing single characters
            {
                string replacement = AbbreviationManager.GetAbbreviation(sel.Text.Trim());
                if (replacement != sel.Text.Trim())
                {
                    sel.Text = replacement; // Replace text in document
                }
            }
        }

        private void Application_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            if (abbreviationEnabled)
            {  // 9 is the key code for TAB
                try
                {
                    Word.Document doc = Application.ActiveDocument;
                    Word.Range selectionRange = doc.Application.Selection.Range;
                    string text = selectionRange.Text.Trim();

                    Application.AutoCorrect.ReplaceText = true; // Enable AutoCorrect
                    selectionRange.Text = abbreviations[text]; // Replace with abbreviation
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show("Error: " + ex.Message, "Status");
                    Debug.WriteLine("Error: " + ex.Message);
                }
            }
        }


        /// <summary>
        /// Enable or disable automatic abbreviation replacement.
        /// </summary>
        public void ToggleAbbreviationReplacement(bool enable)
        {
            abbreviationEnabled = enable;
            if (abbreviationEnabled)
            {
                // Optionally, update the UI or status to indicate the feature is enabled
                System.Windows.Forms.MessageBox.Show("Abbreviation Replacement Enabled", "Status");
            }
            else
            {
                // Optionally, update the UI or status to indicate the feature is disabled
                System.Windows.Forms.MessageBox.Show("Abbreviation Replacement Disabled", "Status");
            }
        }

        /// <summary>
        /// Replace all abbreviations in the active document at once.
        /// </summary>
        public void ReplaceAllAbbreviations()
        {
            Word.Document doc = this.Application.ActiveDocument;
            Word.Range range = doc.Content;
            foreach (var phrase in AbbreviationManager.GetAllPhrases())
            {
                string abbreviation = AbbreviationManager.GetAbbreviation(phrase);
                range.Find.Execute(FindText: phrase, ReplaceWith: abbreviation, Replace: Word.WdReplace.wdReplaceAll);
            }
        }

        public void ReplaceAutoCorrectAbbreviationsInWord()
        {
            var application = Globals.ThisAddIn.Application;
            var activeDocument = application.ActiveDocument;

            try
            {
                // Access the AutoCorrect object
                AutoCorrect autoCorrect = application.AutoCorrect;

                // Create a list of abbreviations and full forms
                var replacements = new List<Tuple<string, string>>();

                // Loop through each AutoCorrect entry
                for (int i = 1; i <= autoCorrect.Entries.Count; i++)
                {
                    string abbreviation = autoCorrect.Entries[i].Name;
                    string fullForm = autoCorrect.Entries[i].Value;

                    // Only add to the list if both abbreviation and full form are not empty
                    if (!string.IsNullOrEmpty(abbreviation) && !string.IsNullOrEmpty(fullForm))
                    {
                        replacements.Add(new Tuple<string, string>(abbreviation, fullForm));
                    }
                }

                // Perform replacements in a single pass using Find.Execute
                foreach (var replacement in replacements)
                {
                    // Clear previous Find settings
                    var find = activeDocument.Content.Find;
                    find.ClearFormatting();
                    find.Replacement.ClearFormatting();
                    find.Text = replacement.Item1;
                    find.Replacement.Text = replacement.Item2;

                    // Execute Find and Replace
                    find.Execute(
                        FindText: replacement.Item1,
                        ReplaceWith: replacement.Item2,
                        Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll
                    );
                }

                System.Windows.Forms.MessageBox.Show("All AutoCorrect abbreviations replaced!", "Success");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error: " + ex.Message, "AutoCorrect Abbreviation Replacement Error");
            }
        }

        private void FindAndReplace(Microsoft.Office.Interop.Word.Document document, string findText, string replaceText)
        {
            // Use the Find and Replace functionality
            var find = document.Content.Find;

            find.ClearFormatting();
            find.Text = findText;
            find.Replacement.ClearFormatting();
            find.Replacement.Text = replaceText;

            // Replace all occurrences in the entire document
            find.Execute(FindText: findText, ReplaceWith: replaceText, Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll);
        }

        public void HighlightAutoCorrectAbbreviationsInWord()
        {
            var application = Globals.ThisAddIn.Application;
            var activeDocument = application.ActiveDocument;

            try
            {
                // Access the AutoCorrect object
                AutoCorrect autoCorrect = application.AutoCorrect;

                // Loop through each AutoCorrect entry
                for (int i = 1; i <= autoCorrect.Entries.Count; i++)
                {
                    string abbreviation = autoCorrect.Entries[i].Name;
                    string fullForm = autoCorrect.Entries[i].Value;

                    // Only highlight if abbreviation is not empty
                    if (!string.IsNullOrEmpty(abbreviation) && !string.IsNullOrEmpty(fullForm))
                    {
                        // Use Find and Highlight for each abbreviation
                        HighlightTextInDocument(activeDocument, abbreviation);
                    }
                }

                System.Windows.Forms.MessageBox.Show("All AutoCorrect abbreviations highlighted!", "Success");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error: " + ex.Message, "AutoCorrect Abbreviation Highlight Error");
            }
        }

        private void HighlightTextInDocument(Microsoft.Office.Interop.Word.Document document, string findText)
        {
            // Use the Find functionality to search for the abbreviation
            var find = document.Content.Find;

            find.ClearFormatting();
            find.Text = findText;

            // Set formatting to highlight found text with a custom color
            find.Replacement.ClearFormatting();
            find.Replacement.Highlight = 1; // Enable Highlight
            find.Replacement.Font.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen; // Green highlight

            // Highlight all occurrences in the entire document
            find.Execute(FindText: findText, ReplaceWith: findText, Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll);
        }

        /// <summary>
        /// Highlight all abbreviations in red in the document.
        /// </summary>
        public void HighlightAllAbbreviations()
        {
            Word.Document doc = this.Application.ActiveDocument;
            Word.Range range = doc.Content;
            foreach (var phrase in AbbreviationManager.GetAllPhrases())
            {
                Word.Find find = range.Find;
                find.ClearFormatting();
                find.Text = phrase;
                find.Replacement.ClearFormatting();
                find.Replacement.Font.Color = Word.WdColor.wdColorRed;
                find.Execute(Replace: Word.WdReplace.wdReplaceAll);
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
