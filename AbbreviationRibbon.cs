using System;
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
            Globals.ThisAddIn.ToggleAbbreviationReplacement(true);

            try
            {

                btnEnable.Enabled = false;  // Disable enable button after click
                btnDisable.Enabled = true;  // Enable disable button
                btnEnable.Label = "Enabled"; // Change text to indicate it's clicked
                                             ////btnEnable.Image = Properties.Resources.enabledIcon; // Optional: Change button icon
                btnDisable.Label = "Disable"; // Reset disable button text


                var application = Globals.ThisAddIn.Application;
                AutoCorrect autoCorrect = application.AutoCorrect;


                autoCorrect.ReplaceText = true; // Disable text replacement
                autoCorrect.CorrectCapsLock = true; // Disable correction of Caps Lock
                autoCorrect.CorrectSentenceCaps = true; // Disable sentence capitalization
                autoCorrect.CorrectInitialCaps = true; // Disable initial capitalization corrections
                autoCorrect.CorrectHangulAndAlphabet = true; // Disable Hangul corrections, if applicable
                autoCorrect.OtherCorrectionsAutoAdd = true; // Disable other corrections auto add




                //System.Windows.Forms.MessageBox.Show("Abbreviator has been enabled and abbreviations added!!", "Success");
            }
            catch (COMException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error: " + ex.Message, "Abbreviator Enabling Error");
            }
        }

        private void btnDisable_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ToggleAbbreviationReplacement(false);

            btnEnable.Enabled = true;  // Enable enable button
            btnDisable.Enabled = false; // Disable disable button
            btnDisable.Label = "Disabled"; // Change text to indicate it's clicked
                                           //btnDisable.Image = Properties.Resources.disabledIcon; // Optional: Change button icon
            btnEnable.Label = "Enable"; // Reset enable button text


            try
            {
                var application = Globals.ThisAddIn.Application;
                AutoCorrect autoCorrect = application.AutoCorrect;

                // Disable various AutoCorrect features

                // Disable various AutoCorrect features
                autoCorrect.ReplaceText = false; // Disable text replacement
                autoCorrect.CorrectCapsLock = false; // Disable correction of Caps Lock
                autoCorrect.CorrectSentenceCaps = false; // Disable sentence capitalization
                autoCorrect.CorrectInitialCaps = false; // Disable initial capitalization corrections
                autoCorrect.CorrectHangulAndAlphabet = false; // Disable Hangul corrections, if applicable
                autoCorrect.OtherCorrectionsAutoAdd = false; // Disable other corrections auto add
            }
            catch (COMException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error: " + ex.Message, "Abbreviator Disabling Error");
            }
        }

        private async void btnReplaceAll_Click(object sender, RibbonControlEventArgs e)
        {
            //Globals.ThisAddIn.ReplaceAllAbbreviations();
            var button = (RibbonButton)sender;
            button.Enabled = false;  // Disable the button
            button.Label = "Processing...";  // Update text to show processing

            //await System.Threading.Tasks.Task.Run(() => Globals.ThisAddIn.ReplaceAutoCorrectAbbreviationsInWord());
            await System.Threading.Tasks.Task.Run(() => Globals.ThisAddIn.ReplaceAllAbbreviations());

            button.Label = "Replace All";  // Reset text
            button.Enabled = true;  // Re-enable the button
        }

        private async void btnHighlightAll_Click(object sender, RibbonControlEventArgs e)
        {
            var button = (RibbonButton)sender;
            button.Enabled = false;  // Disable the button
            button.Label = "Processing...";  // Show processing message

            //await System.Threading.Tasks.Task.Run(() => Globals.ThisAddIn.HighlightAutoCorrectAbbreviationsInWord());
            await System.Threading.Tasks.Task.Run(() => Globals.ThisAddIn.HighlightAllAbbreviations());

            button.Label = "Highlight All";  // Reset label
            button.Enabled = true;  // Re-enable the button
        }
    }
}