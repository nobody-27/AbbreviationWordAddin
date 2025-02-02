﻿using System;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools.Ribbon;

namespace AbbreviationWordAddin
{
    public partial class AbbreviationRibbon
    {
        private void AbbreviationRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            // Check if the abbreviation replacement is enabled
            var application = Globals.ThisAddIn.Application;
            AutoCorrect autoCorrect = application.AutoCorrect;
            if (autoCorrect.ReplaceText)
            {
                // Initialize cache on load if enabled
                AbbreviationManager.InitializeAutoCorrectCache(autoCorrect);
                
                Globals.ThisAddIn.ToggleAbbreviationReplacement(true);
                btnEnable.Enabled = false;  // Disable enable button
                btnDisable.Enabled = true;  // Enable disable button
                btnEnable.Label = "Enabled"; // Change text to indicate it's enabled
                btnDisable.Label = "Disable"; // Reset disable button text
            }
            else
            {
                Globals.ThisAddIn.ToggleAbbreviationReplacement(false);
                btnEnable.Enabled = true;  // Enable enable button
                btnDisable.Enabled = false; // Disable disable button
                btnDisable.Label = "Disabled"; // Change text to indicate it's disabled
                btnEnable.Label = "Enable"; // Reset enable button text
            }
        }

        private void btnEnable_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var application = Globals.ThisAddIn.Application;
                AutoCorrect autoCorrect = application.AutoCorrect;

                // Initialize cache before enabling
                AbbreviationManager.InitializeAutoCorrectCache(autoCorrect);
                
                Globals.ThisAddIn.ToggleAbbreviationReplacement(true);

                btnEnable.Enabled = false;  // Disable enable button after click
                btnDisable.Enabled = true;  // Enable disable button
                btnEnable.Label = "Enabled"; // Change text to indicate it's clicked
                btnDisable.Label = "Disable"; // Reset disable button text

                autoCorrect.ReplaceText = true; // Enable text replacement
                autoCorrect.CorrectCapsLock = true; 
                autoCorrect.CorrectSentenceCaps = true;
                autoCorrect.CorrectInitialCaps = true;
                autoCorrect.CorrectHangulAndAlphabet = true;
                autoCorrect.OtherCorrectionsAutoAdd = true;
            }
            catch (COMException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error: " + ex.Message, "Abbreviator Enabling Error");
            }
        }

        private void btnDisable_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var application = Globals.ThisAddIn.Application;
                AutoCorrect autoCorrect = application.AutoCorrect;

                // Clear cache before disabling
                AbbreviationManager.ClearAutoCorrectCache();
                
                Globals.ThisAddIn.ToggleAbbreviationReplacement(false);

                btnEnable.Enabled = true;  // Enable enable button
                btnDisable.Enabled = false; // Disable disable button
                btnDisable.Label = "Disabled"; // Change text to indicate it's clicked
                btnEnable.Label = "Enable"; // Reset enable button text

                // Disable various AutoCorrect features
                autoCorrect.ReplaceText = false;
                autoCorrect.CorrectCapsLock = false;
                autoCorrect.CorrectSentenceCaps = false;
                autoCorrect.CorrectInitialCaps = false;
                autoCorrect.CorrectHangulAndAlphabet = false;
                autoCorrect.OtherCorrectionsAutoAdd = false;
            }
            catch (COMException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error: " + ex.Message, "Abbreviator Disabling Error");
            }
        }

        private async void btnReplaceAll_Click(object sender, RibbonControlEventArgs e)
        {
            var button = (RibbonButton)sender;
            button.Enabled = false;  // Disable the button
            button.Label = "Processing...";  // Update text to show processing

            try
            {
                // Ensure cache is initialized before processing
                if (!AbbreviationManager.IsAutoCorrectCacheInitialized())
                {
                    AbbreviationManager.InitializeAutoCorrectCache(Globals.ThisAddIn.Application.AutoCorrect);
                }

                await System.Threading.Tasks.Task.Run(() => Globals.ThisAddIn.ReplaceAllAbbreviations());
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error during replacement: " + ex.Message, "Error");
            }
            finally
            {
                button.Label = "Replace All";  // Reset text
                button.Enabled = true;  // Re-enable the button
            }
        }

        private async void btnHighlightAll_Click(object sender, RibbonControlEventArgs e)
        {
            var button = (RibbonButton)sender;
            button.Enabled = false;  // Disable the button
            button.Label = "Processing...";  // Show processing message

            try
            {
                // Ensure cache is initialized before processing
                if (!AbbreviationManager.IsAutoCorrectCacheInitialized())
                {
                    AbbreviationManager.InitializeAutoCorrectCache(Globals.ThisAddIn.Application.AutoCorrect);
                }

                await System.Threading.Tasks.Task.Run(() => Globals.ThisAddIn.HighlightAllAbbreviations());
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error during highlighting: " + ex.Message, "Error");
            }
            finally
            {
                button.Label = "Highlight All";  // Reset label
                button.Enabled = true;  // Re-enable the button
            }
        }
    }
}
