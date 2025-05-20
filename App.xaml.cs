// App.xaml.cs

using System;
using System.Configuration;
using System.Data;
using System.Windows;
using System.Windows.Threading; // For DispatcherUnhandledExceptionEventArgs

namespace AudioFusion
{
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
            string errorMessage = $"An unhandled UI exception occurred:\n\n{e.Exception.Message}\n\nStack Trace:\n{e.Exception.StackTrace}";
            if (e.Exception.InnerException != null)
            {
                errorMessage += $"\n\nInner Exception:\n{e.Exception.InnerException.Message}\n{e.Exception.InnerException.StackTrace}";
            }
            System.Diagnostics.Debug.WriteLine($"UI THREAD ERROR: {errorMessage}");
            MessageBox.Show(errorMessage, "UI Error", MessageBoxButton.OK, MessageBoxImage.Error);
            e.Handled = true; // Prevent application termination for UI thread exceptions if possible
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            var exception = e.ExceptionObject as Exception;
            string errorMessage = $"A F_A_T_A_L non-UI error occurred (IsTerminating: {e.IsTerminating}):\n\n{exception?.Message ?? "Unknown error"}";
            if (exception?.StackTrace != null)
            {
                errorMessage += $"\n\nStack Trace:\n{exception.StackTrace}";
            }
            if (exception?.InnerException != null)
            {
                errorMessage += $"\n\nInner Exception:\n{exception.InnerException.Message}\n{exception.InnerException.StackTrace}";
            }
            System.Diagnostics.Debug.WriteLine($"NON-UI THREAD FATAL ERROR: {errorMessage}");

            // Try to show a message, but the app might be unstable or terminating.
            // Dispatch to UI thread if possible and if Current is available.
            // Use Invoke rather than InvokeAsync for critical errors to ensure it attempts to show before potential termination.
            Current?.Dispatcher?.Invoke(() =>
            {
                try
                {
                    MessageBox.Show(errorMessage, "Fatal Non-UI Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                catch (Exception ex)
                {
                    // If MessageBox fails (e.g., UI thread is gone), log this failure.
                    System.Diagnostics.Debug.WriteLine($"Failed to show MessageBox for non-UI error: {ex.Message}");
                }
            });

            // For truly fatal errors, consider synchronous logging to a file here,
            // as async operations or UI interactions might not complete before termination.
            // Environment.Exit(1); // Optionally, explicitly terminate if e.IsTerminating is true.
        }
    }
}