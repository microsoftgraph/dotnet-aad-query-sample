// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using CommunityToolkit.Mvvm.Messaging;
using MsGraphSamples.WPF.Helpers;

namespace MsGraphSamples.WPF.Views;

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

    // Generate DataGridColumns
    public void Receive(string[] splittedSelect)
    {
        DirectoryObjectsGrid.Columns.Clear();

        foreach (var property in splittedSelect)
        {
            // handle extension properties
            if (property.StartsWith("extension_"))
            {
                DirectoryObjectsGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = property.Split('_')[2],
                    Binding = new Binding("AdditionalData") { Converter = new AdditionalDataConverter(), ConverterParameter = property },
                    Width = new DataGridLength(1, DataGridLengthUnitType.Star)
                });
            }
            else
            {
                // TODO: find a more robust way to generate bindings with property names
                DirectoryObjectsGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = property,
                    // binding needs exact casing of the property name (e.g. "UserPrincipalName" instead of "userPrincipalName")
                    Binding = new Binding(char.ToUpper(property[0]) + property[1..]),
                    Width = new DataGridLength(1, DataGridLengthUnitType.Star)
                });
            }
        }
    }
}