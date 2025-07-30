# AudioFusion

AudioFusion is a Windows desktop application for merging audio outputs and inputs. It uses **NAudio** to capture and route audio streams, providing two core capabilities:

1. **Output Fusion** – Mirror system audio to a second headset and control its volume.
2. **Input Fusion** – Combine the system default microphone with another microphone into a single virtual stream.

The UI is built with WPF and targets **.NET 8.0**.

## Features

- Lists available playback and recording devices.
- Mirrors system playback to a second output device.
- Adjusts volume on the secondary headset.
- Mixes two microphones and exposes the result as a virtual output.
- Refreshes device lists without restarting the application.

## Project Layout

- `App.xaml` / `App.xaml.cs` – Application entry and global exception handling.
- `MainWindow.xaml` / `MainWindow.xaml.cs` – Main UI and all audio routing logic.
- `VirtualAudioCaptureClient.cs` – Provides a simple virtual output for the mixed microphone stream.
- `SampleProvider.cs`, `SampleToWaveProvider.cs`, `WaveExtensionMethods.cs` – Helper classes for converting between `IWaveProvider` and `ISampleProvider`.
- `AudioFusion.csproj` – Project file referencing **NAudio** and **Microsoft.Graph** (graph library is not currently used in code).

## Building

Requirements:

- Windows with the [.NET 8.0 SDK](https://dotnet.microsoft.com/download) installed.
- Visual Studio 2022 or the `dotnet` CLI.

To build from the command line:

```bash
 dotnet restore
 dotnet build AudioFusion.csproj -c Release
```

## Running

After building, run the produced executable (`AudioFusion.exe`). The interface lets you choose audio devices and start or stop the fusion features.

## License

This project is provided for **non‑commercial use only**. You may use, modify and distribute it freely for personal or educational purposes. Commercial usage of any kind is not permitted without prior written consent.

