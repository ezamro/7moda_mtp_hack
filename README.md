# 7moda MTP Tool

A lightweight Windows tool for browsing and transferring files from Android phones and other MTP devices. No drivers or admin rights needed — uses the built-in Windows MTP API.

## Features

- **List devices** — Detect all connected MTP phones/tablets/cameras
- **Device info** — Show model, manufacturer, serial, battery level, storage
- **Browse files** — List files and folders with sizes and dates
- **Folder tree** — Visual tree view of the entire device storage
- **Download files** — Copy files from device to PC

## Usage

```
7moda list                    — List connected MTP devices
7moda info <index|name>       — Show detailed device info
7moda dir <index|name> [path] — List files on device
7moda tree <index|name> [p]   — Show full folder tree
7moda get <idx> <src> <dst>   — Copy file from device
```

### Examples

```
7moda list
7moda info 0
7moda dir 0 "DCIM/Camera"
7moda tree 0
7moda get 0 "DCIM/photo.jpg" C:\temp\photo.jpg
```

## Requirements

- Windows 7 or later
- Phone connected via USB in MTP (File Transfer) mode
- Phone screen unlocked

No installation required. Just download and run `7moda.exe`.

## Build from Source

```cmd
git clone <repo-url>
cd 7moda-mtp-tool
dotnet publish -o dist
```

Requires .NET 8 SDK.

## License

MIT
