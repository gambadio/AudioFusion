# 🎵 AudioFusion

**Professional Audio Routing & Mixing for Windows**

AudioFusion is a powerful desktop application that revolutionizes audio management on Windows by providing advanced audio device fusion, routing, and real-time mixing capabilities. Perfect for content creators, streamers, musicians, and audio professionals who need seamless control over multiple audio sources.

## ✨ Core Features

### 🎛️ **Output Fusion**
- **Multi-Device Audio Routing**: Route any application's audio to multiple output devices simultaneously
- **Real-Time Audio Mirroring**: Send system audio to headphones while broadcasting to speakers
- **Zero-Latency Processing**: Optimized for real-time audio with minimal delay
- **Volume Independent Control**: Separate volume control for each output destination

### 🎤 **Input Fusion** 
- **Multi-Microphone Mixing**: Combine multiple microphones into a single virtual input
- **Dynamic Audio Blending**: Real-time mixing with individual volume controls
- **Virtual Microphone Creation**: Create a virtual audio device for streaming applications
- **Professional Audio Quality**: High-fidelity audio processing throughout the chain

### 🔧 **Advanced Audio Control**
- **Device Management**: Comprehensive control over all audio input and output devices
- **Real-Time Monitoring**: Live audio level visualization and monitoring
- **Flexible Routing Options**: Customize audio paths for any workflow
- **Low-Latency WASAPI Integration**: Utilizes Windows Audio Session API for optimal performance

## 🚀 Getting Started

### System Requirements

- **Operating System**: Windows 10 or later
- **Framework**: .NET Framework 4.7.2 or higher
- **Audio**: WASAPI-compatible audio devices
- **Memory**: 4GB RAM minimum, 8GB recommended for complex routing

### Installation

1. **Download the latest release** from the GitHub releases page
2. **Extract the application** to your preferred directory
3. **Run AudioFusion.exe** as administrator (required for audio device access)
4. **Configure your audio devices** through the intuitive interface

### Quick Setup

1. **Launch AudioFusion** and grant audio device permissions
2. **Select Primary Devices**: Choose your main input and output devices
3. **Configure Fusion Settings**: Set up output or input fusion based on your needs
4. **Start Fusion**: Click the appropriate start button to begin audio routing
5. **Adjust Levels**: Use the volume sliders to balance audio sources

## 📖 Usage Guide

### Output Fusion Workflow
1. **Select Audio Source**: Choose the application or system audio to route
2. **Choose Primary Output**: Set your main audio output device (e.g., speakers)
3. **Add Secondary Output**: Select additional output device (e.g., headphones)
4. **Configure Volume**: Adjust independent volume levels for each output
5. **Start Fusion**: Begin real-time audio routing and enjoy simultaneous output

### Input Fusion Workflow  
1. **Connect Multiple Microphones**: Ensure all desired microphones are connected
2. **Select Primary Microphone**: Choose your main audio input source
3. **Add Secondary Sources**: Select additional microphones to mix
4. **Balance Audio Levels**: Adjust individual microphone volumes
5. **Create Virtual Input**: Generate a combined virtual microphone for applications

### Professional Tips
- **Monitor Audio Levels**: Keep an eye on the real-time level indicators to prevent clipping
- **Optimize Buffer Settings**: Adjust audio buffer sizes for your specific hardware configuration
- **Use Exclusive Mode**: Enable exclusive audio mode for lowest possible latency
- **Test Before Recording**: Always test your fusion setup before important recordings or streams

## 🏗️ Technical Architecture

### Core Components

**WASAPI Integration**
- Low-level Windows Audio Session API integration
- Exclusive and shared mode operation
- Real-time audio capture and playback
- Professional-grade audio buffer management

**NAudio Framework**
- Comprehensive .NET audio library utilization
- Advanced sample providers and mixing capabilities
- Multi-format audio processing support
- Robust error handling and recovery

**WPF User Interface**
- Modern, responsive Windows Presentation Foundation interface
- Real-time audio level visualization
- Intuitive device management controls
- Professional audio application design patterns

### Dependencies
- **NAudio**: Professional audio library for .NET applications
- **NAudio.CoreAudioApi**: Core Audio API bindings for Windows
- **NAudio.Wave**: Wave audio processing and manipulation
- **.NET Framework**: Microsoft's application development platform

## ⚙️ Configuration

### Audio Device Settings
- **Sample Rate**: Configure sample rates (44.1kHz, 48kHz, 96kHz)
- **Bit Depth**: Select bit depth (16-bit, 24-bit, 32-bit)
- **Buffer Size**: Adjust buffer sizes for latency vs. stability balance
- **Exclusive Mode**: Enable for professional low-latency operation

### Fusion Options
- **Output Routing**: Customize how audio is distributed to multiple outputs
- **Input Mixing**: Configure microphone blending and virtual device creation
- **Volume Mapping**: Set up automatic gain control and level management
- **Device Monitoring**: Configure real-time audio level displays

## 🛠️ Troubleshooting

### Common Issues

**"Audio Device Not Found" Errors**
- Verify all audio devices are properly connected and recognized by Windows
- Run AudioFusion as Administrator to ensure device access permissions
- Check Windows Sound settings for device availability and default assignments

**High Latency or Audio Dropouts**
- Reduce audio buffer sizes in the configuration settings
- Close unnecessary applications that may be using audio resources
- Enable exclusive mode for professional audio interfaces
- Ensure sufficient system resources are available

**Virtual Microphone Not Appearing**
- Restart audio applications after creating virtual input devices
- Check Windows Sound settings for the virtual device recognition
- Verify that input fusion is properly started and configured

**Volume Control Issues**
- Check both application and system volume levels
- Verify that audio devices support the requested volume ranges
- Test with different audio sources to isolate the issue

## 🎯 Use Cases

### Content Creation
- **Streaming**: Route game audio to headphones while sending commentary to streaming software
- **Podcasting**: Mix multiple microphones for professional multi-host recordings
- **Video Production**: Separate audio tracks for post-production flexibility

### Music Production
- **Studio Monitoring**: Send audio to multiple monitoring systems simultaneously
- **Live Performance**: Route instruments to different outputs for stage and recording
- **Audio Engineering**: Professional multi-channel audio routing and mixing

### Gaming & Communication
- **Team Gaming**: Balance game audio with voice chat for optimal communication
- **Content Streaming**: Professional audio setup for gaming broadcasts
- **Discord/Teams**: Enhanced audio quality for voice communication applications

## 🤝 Contributing

AudioFusion welcomes contributions to enhance audio processing capabilities:

- **Feature Development**: Additional audio effects, routing options, and device support
- **UI/UX Improvements**: Enhanced user interface and workflow optimization
- **Performance Optimization**: Audio processing efficiency and latency improvements  
- **Device Compatibility**: Extended support for additional audio hardware
- **Documentation**: User guides, tutorials, and technical documentation

## 📄 License

This project is licensed for **non-commercial use only**.

**Personal Use**: Free to use, copy, and modify for private, educational, and non-commercial purposes.  
**Commercial Use**: Prohibited without explicit written permission.

For commercial licensing inquiries, please contact: **ricardo.kupper@adalala.com**

## 🌟 Acknowledgments

- **NAudio Project**: For providing the comprehensive .NET audio framework
- **Microsoft WASAPI**: For low-latency audio capabilities on Windows
- **Audio Community**: For inspiration and feedback that shaped this professional tool
- **Beta Testers**: For their invaluable testing and feature suggestions

---

*Professional audio routing made simple - where multiple audio sources become one seamless experience.*

**© 2025 Ricardo Kupper** • Built with precision for audio professionals and enthusiasts
