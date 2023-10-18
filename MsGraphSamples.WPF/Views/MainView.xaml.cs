﻿// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using MsGraph_Samples.ViewModels;

namespace MsGraph_Samples.Views;

public partial class MainView : Window
{
    private MainViewModel ViewModel => (MainViewModel)DataContext;

    public MainView()
    {
        InitializeComponent();
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

    private void ResultsDataGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
    {
        if (!ViewModel.SplittedSelect.Any())
            return;

        e.Cancel = !e.PropertyName.In(ViewModel.SplittedSelect);
        if (!e.Cancel)
        {
            e.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
        }
    }

    private void ResultsDataGrid_AutoGeneratedColumns(object sender, System.EventArgs e)
    {
        var dg = (DataGrid)sender;
        foreach (var column in dg.Columns)
        {
            column.DisplayIndex = Array.FindIndex(
                ViewModel.SplittedSelect,
                p => p.Equals(column.Header.ToString(), StringComparison.OrdinalIgnoreCase));
        }
    }

}