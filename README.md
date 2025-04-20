# ucAniGifEx 🎉

Welcome to the **ucAniGifEx** repository! This project provides an ActiveX control for handling animated GIFs seamlessly in your applications. With this control, you can easily incorporate dynamic GIF animations into your projects, enhancing user experience and visual appeal.

![ucAniGifEx](https://img.shields.io/badge/Download%20Latest%20Release-Click%20Here-brightgreen?style=flat-square&logo=github)

## Table of Contents

- [Introduction](#introduction)
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Examples](#examples)
- [Contributing](#contributing)
- [License](#license)
- [Contact](#contact)

## Introduction

The **ucAniGifEx** control is designed for developers who need to display animated GIFs in their applications. This control is built using the Windows Imaging Component (WIC) and Direct2D, ensuring high-quality rendering of GIF animations. Whether you are working with VB6, VBA, or TwinBasic, this control simplifies the process of incorporating GIFs into your projects.

## Features

- **Easy Integration**: Quickly add animated GIFs to your applications with minimal setup.
- **Supports Multiple GIF Formats**: Handle various GIF animations without issues.
- **Optimized Performance**: Utilizes Direct2D for smooth rendering and playback.
- **User-Friendly Interface**: Simple properties and methods for easy manipulation.
- **Cross-Platform Compatibility**: Works with various Windows platforms.

## Installation

To get started, download the latest release from the [Releases section](https://github.com/Nessanggg/ucAniGifEx/releases). After downloading, follow these steps:

1. Extract the files from the downloaded package.
2. Register the ActiveX control using the command:
   ```bash
   regsvr32 ucAniGifEx.ocx
   ```
3. Add the control to your project by selecting it from the toolbox.

For more detailed instructions, please refer to the documentation included in the release package.

## Usage

Using the **ucAniGifEx** control is straightforward. Here’s a simple example to help you get started:

1. Drag and drop the **ucAniGifEx** control onto your form.
2. Set the `GIFPath` property to the path of your GIF file:
   ```vb
   ucAniGifEx1.GIFPath = "C:\path\to\your\animation.gif"
   ```
3. Call the `Play` method to start the animation:
   ```vb
   ucAniGifEx1.Play
   ```

You can stop the animation using the `Stop` method:
```vb
ucAniGifEx1.Stop
```

For more advanced features, check the documentation provided in the release package.

## Examples

Here are a few examples of how to use the **ucAniGifEx** control in different scenarios:

### Example 1: Basic Animation

```vb
Private Sub Form_Load()
    ucAniGifEx1.GIFPath = "C:\path\to\your\animation.gif"
    ucAniGifEx1.Play
End Sub
```

### Example 2: Looping Animation

```vb
Private Sub Form_Load()
    ucAniGifEx1.GIFPath = "C:\path\to\your\looping_animation.gif"
    ucAniGifEx1.Loop = True
    ucAniGifEx1.Play
End Sub
```

### Example 3: Stopping Animation on Button Click

```vb
Private Sub btnStop_Click()
    ucAniGifEx1.Stop
End Sub
```

## Contributing

We welcome contributions to enhance the **ucAniGifEx** project. If you would like to contribute, please follow these steps:

1. Fork the repository.
2. Create a new branch for your feature or bug fix.
3. Make your changes and commit them with clear messages.
4. Push your changes to your forked repository.
5. Submit a pull request.

Your contributions help improve the control for everyone.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contact

For any questions or support, feel free to reach out:

- **Email**: support@example.com
- **GitHub**: [Nessanggg](https://github.com/Nessanggg)

To download the latest version of the control, visit the [Releases section](https://github.com/Nessanggg/ucAniGifEx/releases). Enjoy using **ucAniGifEx** and happy coding!