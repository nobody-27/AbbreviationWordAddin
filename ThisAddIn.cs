﻿﻿﻿﻿﻿﻿﻿using System;
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
        private const int CHUNK_SIZE = 1000; // Process 1000 words at a time

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WindowSelectionChange += Application_WindowSelectionChange;

            try
            {
                // Load abbreviations from cache or Excel
                AbbreviationManager.LoadAbbreviations();

                var application = Globals.ThisAddIn.Application;
                AutoCorrect autoCorrect = application.AutoCorrect;
                
                // Set initial state based on AutoCorrect settings
                abbreviationEnabled = autoCorrect.ReplaceText;

                // Initialize AutoCorrect cache
                AbbreviationManager.InitializeAutoCorrectCache(autoCorrect);

                // Add each abbreviation to AutoCorrect if enabled
                if (abbreviationEnabled)
                {
                    foreach (var abbreviation in AbbreviationManager.GetAllPhrases())
                    {
                        try
                        {
                            string fullForm = AbbreviationManager.GetAbbreviation(abbreviation);
                            if (!string.IsNullOrEmpty(fullForm))
                            {
                                autoCorrect.ReplaceText = true;
                                autoCorrect.Entries.Add(abbreviation, fullForm);
                            }
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            // Skip if entry already exists
                            continue;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Error during startup: " + ex.Message,
                    "Startup Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning
                );
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Clear cache on shutdown
            AbbreviationManager.ClearAutoCorrectCache();
        }

        /// <summary>
        /// Automatically replaces text when the user types if abbreviations are enabled.
        /// </summary>
        private void Application_WindowSelectionChange(Word.Selection sel)
        {
            if (abbreviationEnabled && sel != null && sel.Text.Length > 3)  // Avoid replacing single characters
            {
                // Try cache first
                string replacement = AbbreviationManager.GetFromAutoCorrectCache(sel.Text.Trim());
                if (replacement == null)
                {
                    replacement = AbbreviationManager.GetAbbreviation(sel.Text.Trim());
                }
                
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

                    // Try cache first
                    string replacement = AbbreviationManager.GetFromAutoCorrectCache(text);
                    if (replacement != null)
                    {
                        Application.AutoCorrect.ReplaceText = true;
                        selectionRange.Text = replacement;
                    }
                    else
                    {
                        string abbrev = AbbreviationManager.GetAbbreviation(text);
                        if (abbrev != text) // Only replace if there's a match
                        {
                            Application.AutoCorrect.ReplaceText = true;
                            selectionRange.Text = abbrev;
                        }
                    }
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
                // Initialize cache when enabling
                AbbreviationManager.InitializeAutoCorrectCache(this.Application.AutoCorrect);
                System.Windows.Forms.MessageBox.Show("Abbreviation Replacement Enabled", "Status");
            }
            else
            {
                // Clear cache when disabling
                AbbreviationManager.ClearAutoCorrectCache();
                System.Windows.Forms.MessageBox.Show("Abbreviation Replacement Disabled", "Status");
            }
        }

        /// <summary>
        /// Replace all abbreviations in the active document at once using optimized chunk processing.
        /// </summary>
        public void ReplaceAllAbbreviations()
        {
            Word.Document doc = this.Application.ActiveDocument;
            
            try
            {
                // Initialize AutoCorrect cache if needed
                if (!AbbreviationManager.IsAutoCorrectCacheInitialized())
                {
                    AbbreviationManager.InitializeAutoCorrectCache(this.Application.AutoCorrect);
                }

                // Get total words in document
                int totalWords = doc.Words.Count;
                
                // Process document in chunks
                for (int startIndex = 1; startIndex <= totalWords; startIndex += CHUNK_SIZE)
                {
                    int endIndex = Math.Min(startIndex + CHUNK_SIZE - 1, totalWords);
                    Word.Range chunkRange = doc.Range(doc.Words[startIndex].Start, doc.Words[endIndex].End);
                    
                    // Store the chunk text
                    string chunkText = chunkRange.Text;
                    bool hasMatches = false;

                    // Quick check if chunk contains any potential matches
                    foreach (var phrase in AbbreviationManager.GetAllPhrases())
                    {
                        if (chunkText.Contains(phrase))
                        {
                            hasMatches = true;
                            break;
                        }
                    }

                    // Only process chunk if it contains potential matches
                    if (hasMatches)
                    {
                        foreach (var phrase in AbbreviationManager.GetAllPhrases())
                        {
                            // Try to get from cache first
                            string replacement = AbbreviationManager.GetFromAutoCorrectCache(phrase);
                            if (replacement == null)
                            {
                                replacement = AbbreviationManager.GetAbbreviation(phrase);
                            }

                            if (chunkText.Contains(phrase))
                            {
                                var find = chunkRange.Find;
                                find.ClearFormatting();
                                find.Text = phrase;
                                find.Replacement.ClearFormatting();
                                find.Replacement.Text = replacement;
                                find.Execute(Replace: Word.WdReplace.wdReplaceAll);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error during replacement: " + ex.Message, "Error");
            }
        }

        /// <summary>
        /// Highlight all abbreviations in the document using optimized chunk processing.
        /// </summary>
        public void HighlightAllAbbreviations()
        {
            Word.Document doc = this.Application.ActiveDocument;
            
            try
            {
                // Initialize AutoCorrect cache if needed
                if (!AbbreviationManager.IsAutoCorrectCacheInitialized())
                {
                    AbbreviationManager.InitializeAutoCorrectCache(this.Application.AutoCorrect);
                }

                // Get total words in document
                int totalWords = doc.Words.Count;
                
                // Process document in chunks
                for (int startIndex = 1; startIndex <= totalWords; startIndex += CHUNK_SIZE)
                {
                    int endIndex = Math.Min(startIndex + CHUNK_SIZE - 1, totalWords);
                    Word.Range chunkRange = doc.Range(doc.Words[startIndex].Start, doc.Words[endIndex].End);
                    
                    // Store the chunk text
                    string chunkText = chunkRange.Text;
                    bool hasMatches = false;

                    // Quick check if chunk contains any potential matches
                    foreach (var phrase in AbbreviationManager.GetAllPhrases())
                    {
                        if (chunkText.Contains(phrase))
                        {
                            hasMatches = true;
                            break;
                        }
                    }

                    // Only process chunk if it contains potential matches
                    if (hasMatches)
                    {
                        foreach (var phrase in AbbreviationManager.GetAllPhrases())
                        {
                            if (chunkText.Contains(phrase))
                            {
                                var find = chunkRange.Find;
                                find.ClearFormatting();
                                find.Text = phrase;
                                find.Replacement.ClearFormatting();
                                find.Replacement.Font.Color = Word.WdColor.wdColorRed;
                                find.Execute(Replace: Word.WdReplace.wdReplaceAll);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error during highlighting: " + ex.Message, "Error");
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
