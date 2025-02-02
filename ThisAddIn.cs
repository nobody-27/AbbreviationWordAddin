﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using Action = System.Action; // Explicitly use System.Action to resolve ambiguity

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
            var progressForm = new ProgressForm();
            var syncContext = System.Threading.SynchronizationContext.Current;
            bool completed = false;
            Exception processError = null;

            // Create a background thread for progress updates
            var progressThread = new System.Threading.Thread(() =>
            {
                try
                {
                    Word.Document doc = null;
                    syncContext.Send(_ =>
                    {
                        // Get document reference on UI thread
                        doc = this.Application.ActiveDocument;
                    }, null);

                    // Initialize AutoCorrect cache if needed
                    if (!AbbreviationManager.IsAutoCorrectCacheInitialized())
                    {
                        syncContext.Send(_ =>
                        {
                            AbbreviationManager.InitializeAutoCorrectCache(this.Application.AutoCorrect);
                        }, null);
                    }

                    int totalWords = 0;
                    syncContext.Send(_ =>
                    {
                        totalWords = doc.Words.Count;
                    }, null);

                    int totalChunks = (totalWords + CHUNK_SIZE - 1) / CHUNK_SIZE;
                    int currentChunk = 0;

                    // Process document in chunks
                    for (int startIndex = 1; startIndex <= totalWords && !completed; startIndex += CHUNK_SIZE)
                    {
                        currentChunk++;
                        int endIndex = Math.Min(startIndex + CHUNK_SIZE - 1, totalWords);
                        
                        // Update progress
                        int percentage = (currentChunk * 100) / totalChunks;
                        progressForm.UpdateProgress(percentage, $"Processing chunk {currentChunk} of {totalChunks}...");

                        // Process chunk on UI thread
                        syncContext.Send(_ =>
                        {
                            try
                            {
                                Word.Range chunkRange = doc.Range(doc.Words[startIndex].Start, doc.Words[endIndex].End);
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

                                if (hasMatches)
                                {
                                    foreach (var phrase in AbbreviationManager.GetAllPhrases())
                                    {
                                        string replacement = AbbreviationManager.GetFromAutoCorrectCache(phrase) 
                                            ?? AbbreviationManager.GetAbbreviation(phrase);

                                        if (chunkText.Contains(phrase))
                                        {
                                            var find = chunkRange.Find;
                                            find.ClearFormatting();
                                            find.Text = phrase;
                                            find.Forward = true;
                                            find.Format = false;
                                            find.MatchCase = false;
                                            find.MatchWholeWord = true;
                                            find.MatchWildcards = false;
                                            find.MatchSoundsLike = false;
                                            find.MatchAllWordForms = false;
                                            find.Wrap = Word.WdFindWrap.wdFindContinue;

                                            find.Replacement.ClearFormatting();
                                            find.Replacement.Text = replacement;

                                            // Execute replacement
                                            find.Execute(
                                                FindText: phrase,
                                                MatchCase: false,
                                                MatchWholeWord: true,
                                                MatchWildcards: false,
                                                MatchSoundsLike: false,
                                                MatchAllWordForms: false,
                                                Forward: true,
                                                Wrap: Word.WdFindWrap.wdFindContinue,
                                                Format: false,
                                                ReplaceWith: replacement,
                                                Replace: Word.WdReplace.wdReplaceAll
                                            );
                                        }
                                    }
                                }

                                // Release COM objects
                                if (chunkRange != null)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(chunkRange);
                            }
                            catch (Exception ex)
                            {
                                processError = ex;
                                completed = true; // Stop processing on error
                            }
                        }, null);
                    }
                }
                catch (Exception ex)
                {
                    processError = ex;
                }
                finally
                {
                    completed = true;
                    syncContext.Post(_ => progressForm.Close(), null);
                }
            });

            progressThread.Start();
            progressForm.ShowDialog();

            if (processError != null)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Error during replacement: " + processError.Message,
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error
                );
            }
        }

        /// <summary>
        /// Highlight all abbreviations in the document using optimized chunk processing.
        /// </summary>
        public void HighlightAllAbbreviations()
        {
            var progressForm = new ProgressForm();
            var syncContext = System.Threading.SynchronizationContext.Current;
            bool completed = false;
            Exception processError = null;

            // Create a background thread for progress updates
            var progressThread = new System.Threading.Thread(() =>
            {
                try
                {
                    Word.Document doc = null;
                    syncContext.Send(_ =>
                    {
                        // Get document reference on UI thread
                        doc = this.Application.ActiveDocument;
                    }, null);

                    // Initialize AutoCorrect cache if needed
                    if (!AbbreviationManager.IsAutoCorrectCacheInitialized())
                    {
                        syncContext.Send(_ =>
                        {
                            AbbreviationManager.InitializeAutoCorrectCache(this.Application.AutoCorrect);
                        }, null);
                    }

                    int totalWords = 0;
                    syncContext.Send(_ =>
                    {
                        totalWords = doc.Words.Count;
                    }, null);

                    int totalChunks = (totalWords + CHUNK_SIZE - 1) / CHUNK_SIZE;
                    int currentChunk = 0;

                    // Process document in chunks
                    for (int startIndex = 1; startIndex <= totalWords && !completed; startIndex += CHUNK_SIZE)
                    {
                        currentChunk++;
                        int endIndex = Math.Min(startIndex + CHUNK_SIZE - 1, totalWords);
                        
                        // Update progress
                        int percentage = (currentChunk * 100) / totalChunks;
                        progressForm.UpdateProgress(percentage, $"Processing chunk {currentChunk} of {totalChunks}...");

                        // Process chunk on UI thread
                        syncContext.Send(_ =>
                        {
                            try
                            {
                                Word.Range chunkRange = doc.Range(doc.Words[startIndex].Start, doc.Words[endIndex].End);
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

                                if (hasMatches)
                                {
                                    foreach (var phrase in AbbreviationManager.GetAllPhrases())
                                    {
                                        if (chunkText.Contains(phrase))
                                        {
                                            var find = chunkRange.Find;
                                            find.ClearFormatting();
                                            find.Text = phrase;
                                            find.Forward = true;
                                            find.Format = true;
                                            find.MatchCase = false;
                                            find.MatchWholeWord = true;
                                            find.MatchWildcards = false;
                                            find.MatchSoundsLike = false;
                                            find.MatchAllWordForms = false;
                                            find.Wrap = Word.WdFindWrap.wdFindContinue;

                                            find.Replacement.ClearFormatting();
                                            find.Replacement.Font.Color = Word.WdColor.wdColorRed;
                                            find.Replacement.Text = phrase;  // Keep the same text, just change color

                                            // Execute highlighting
                                            find.Execute(
                                                FindText: phrase,
                                                MatchCase: false,
                                                MatchWholeWord: true,
                                                MatchWildcards: false,
                                                MatchSoundsLike: false,
                                                MatchAllWordForms: false,
                                                Forward: true,
                                                Wrap: Word.WdFindWrap.wdFindContinue,
                                                Format: true,
                                                ReplaceWith: phrase,
                                                Replace: Word.WdReplace.wdReplaceAll
                                            );
                                        }
                                    }
                                }

                                // Release COM objects
                                if (chunkRange != null)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(chunkRange);
                            }
                            catch (Exception ex)
                            {
                                processError = ex;
                                completed = true; // Stop processing on error
                            }
                        }, null);
                    }
                }
                catch (Exception ex)
                {
                    processError = ex;
                }
                finally
                {
                    completed = true;
                    syncContext.Post(_ => progressForm.Close(), null);
                }
            });

            progressThread.Start();
            progressForm.ShowDialog();

            if (processError != null)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Error during highlighting: " + processError.Message,
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error
                );
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
