using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddInCustomRibbon
{
    public partial class RibbonCustomJota
    {
        private void RibbonCustomJota_Load(object sender, RibbonUIEventArgs e)
        {

        }
        private void gallerySymbols_Load(object sender, RibbonControlEventArgs e)
        {
            gallerySymbols.Items.Clear();
            // Add symbols to the Gallery
            AddSymbolToGallery("\u00B6", "Paragraph Symbol"); // ¶
            AddSymbolToGallery("\u00A7", "Section Symbol"); // §
            AddSymbolToGallery("\u00A9", "Copyright Symbol"); // ©
            AddSymbolToGallery("\u2122", "Trademark Symbol"); // ™
            AddSymbolToGallery("\u20AC", "Euro Symbol"); // €
            AddSymbolToGallery("\u00BA", "Masculine Ordinal Indicator"); // º
            AddSymbolToGallery("\u00A2", "Cent Symbol"); // ¢
        }
        private void AddSymbolToGallery(string symbol, string label)
        {

            RibbonDropDownItem itemToAdd = Factory.CreateRibbonDropDownItem();
            itemToAdd.Label = string.Empty;
            itemToAdd.Tag = symbol;

            itemToAdd.Image = GetImageForSymbol(symbol);
            gallerySymbols.Items.Add(itemToAdd);
        }

        private Bitmap GetImageForSymbol(string symbol)
        {
            // Create an image with the symbol using a font
            System.Drawing.Font font = new System.Drawing.Font("Arial", 14, FontStyle.Regular, GraphicsUnit.Pixel);
            Bitmap image = new Bitmap(13, 19);
            using (Graphics g = Graphics.FromImage(image))
            {
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                // Measure the size of the symbol
                SizeF textSize = g.MeasureString(symbol, font);
                // Calculate the position to center the symbol
                float x = (image.Width - textSize.Width) / 2;
                float y = (image.Height - textSize.Height) / 2;

                // Adjust x position to center the symbol horizontally
                x += 2.8f;

                g.DrawString(symbol, font, Brushes.Black, new PointF(x, y));
            }
            return image;
        }

        private void gallerySymbols_Click(object sender, RibbonControlEventArgs e)
        {
            // Handle selection change event
            RibbonDropDownItem selectedItem = gallerySymbols.SelectedItem;
            if (selectedItem != null)
            {
                string selectedSymbol = selectedItem.Tag as string;
                InsertSymbolIntoDocument(selectedSymbol);
            }
        }

        private void InsertSymbolIntoDocument(string symbol)
        {
            try
            {
                Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
                Word.Selection selection = Globals.ThisAddIn.Application.Selection;

                if (document != null && selection != null)
                {
                    // Check if the current selection is an image
                    if (selection.Type == Word.WdSelectionType.wdSelectionInlineShape ||
                        selection.Type == Word.WdSelectionType.wdSelectionShape)
                    {
                        // Handle the case when an image is selected
                        HandleImageSelection();
                    }
                    else
                    {
                        // Insert the symbol at the current cursor position
                        selection.TypeText(symbol);
                    }
                }
                else
                {
                    // Handle the case when there is no active document or selection
                    HandleNoDocumentOrSelection();
                }
            }
            catch (Exception ex)
            {
                // Handle any other exceptions that might occur
                HandleException(ex.Message, "Inserting symbol");
            }
        }
        private void HandleImageSelection()
        {
            // Handle the case when an image is selected
            // You can add your custom logic here
            System.Windows.Forms.MessageBox.Show("Cannot insert symbol. An image is selected.");
        }
        private void HandleNoDocumentOrSelection()
        {
            // Handle the case when there is no active document or selection
            // You can add your custom logic here
            System.Windows.Forms.MessageBox.Show("Cannot insert symbol. No active document or selection.");
        }

        private void HandleException(string errorMessage, string eventName)
        {
            // Handle other exceptions
            // You can add your custom logic here
            System.Windows.Forms.MessageBox.Show($"{eventName}: An error occurred, please share this error message with IT: {errorMessage}");
        }

        private void btnAddSymbols_Click(object sender, RibbonControlEventArgs e)
        {
            OpenCharacterMap();
        }
        private void OpenCharacterMap()
        {
            try
            {
                // Launch Character Map
                Process.Start("charmap.exe");
            }
            catch (Exception ex)
            {
                // Handle any exceptions
                HandleException(ex.Message, "Adding new symbol");
            }
        }
    }
}
