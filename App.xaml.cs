using System;
using System.Configuration;
using System.Data;
using System.Windows;
using System.Windows.Threading;

namespace AudioFusion;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    public App()
    {
        // Add handler for unhandled exceptions
        this.DispatcherUnhandledException += App_DispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
    }

    private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        string errorMessage = $"An unhandled exception occurred:\n\n{e.Exception.Message}\n\nStack Trace:\n{e.Exception.StackTrace}";
        MessageBox.Show(errorMessage, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        e.Handled = true;
    }

    private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        var exception = e.ExceptionObject as Exception;
        string errorMessage = $"A fatal error occurred:\n\n{exception?.Message ?? "Unknown error"}\n\nStack Trace:\n{exception?.StackTrace}";
        MessageBox.Show(errorMessage, "Fatal Error", MessageBoxButton.OK, MessageBoxImage.Error);
    }
}

