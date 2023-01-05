// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using CommunityToolkit.Mvvm.Messaging;

namespace MsGraph_Samples.Views
{
    public partial class MainView : Window, IRecipient<string[]>
    {
        public MainView()
        {
            InitializeComponent();
            WeakReferenceMessenger.Default.Register(this);
        }

        private void TextBox_SelectAll(object sender, RoutedEventArgs e)
        {
            var textBox = (TextBox)sender;
            textBox.SelectAll();
        }

        private void TextBox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            var textBox = (TextBox)sender;
            if (!textBox.IsKeyboardFocusWithin)
            {
                textBox.Focus();
                e.Handled = true;
            }
        }

        private void LoadButton_Click(object sender, RoutedEventArgs e)
        {
            // Workaround to prevent IsDefault button to execute before TextBox Bindings
            var button = (Button)sender;
            button.Focus();
        }

        // TODO: find a more robust way
        public void Receive(string[] splittedSelect)
        {
            DirectoryObjectsGrid.Columns.Clear();

            foreach (var property in splittedSelect)
            {
                var column = new DataGridTextColumn
                {
                    Header = property,
                    // binding needs exact casing of the property name
                    Binding = new Binding(char.ToUpper(property[0]) + property[1..]),
                    Width = new DataGridLength(1, DataGridLengthUnitType.Star)
                };
                
                DirectoryObjectsGrid.Columns.Add(column);
            }
        }
    }
}