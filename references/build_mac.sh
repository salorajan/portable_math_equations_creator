#!/bin/bash
# ===================================================
#             BUILDING CONVERT_MD FOR MACOS
# ===================================================

echo "Detecting system architecture..."
ARCH=$(uname -m)
echo "Architecture detected: $ARCH"

# Create a temporary workspace for binary staging
mkdir -p bin_temp

# Select appropriate Pandoc and Typst binaries based on macOS architecture
if [ "$ARCH" = "arm64" ]; then
    echo "Staging Apple Silicon (M-series) binaries..."
    cp bin/pandoc-mac-arm64 bin_temp/pandoc
    cp bin/typst-mac-mac-arm64 bin_temp/typst 2>/dev/null || cp bin/typst-mac-arm64 bin_temp/typst
elif [ "$ARCH" = "x86_64" ]; then
    echo "Staging Intel x86_64 binaries..."
    cp bin/pandoc-mac-x86_64 bin_temp/pandoc
    cp bin/typst-mac-x86_64 bin_temp/typst
else
    echo "Warning: Unknown architecture $ARCH. Defaulting to Intel binaries."
    cp bin/pandoc-mac-x86_64 bin_temp/pandoc
    cp bin/typst-mac-x86_64 bin_temp/typst
fi

# Ensure compiled staged binaries are executable
chmod +x bin_temp/pandoc
chmod +x bin_temp/typst

# Verify PyInstaller is installed
if ! command -v pyinstaller &> /dev/null; then
    echo "[ERROR] PyInstaller is not installed or not in PATH."
    echo "Please install it using: pip install pyinstaller"
    rm -rf bin_temp
    exit 1
fi

echo "Packaging convert_md.py for macOS..."
pyinstaller --onefile \
    --add-data "bin_temp/pandoc:bin" \
    --add-data "bin_temp/typst:bin" \
    convert_md.py

if [ $? -eq 0 ]; then
    echo ""
    echo "[SUCCESS] Build completed! Mac executable is located in dist/convert_md"
    # Copy to bin directory as convert_md_mac for multiplatform distribution
    cp dist/convert_md bin/convert_md_mac
    chmod +x bin/convert_md_mac
    echo "Copied executable to bin/convert_md_mac"
else
    echo ""
    echo "[ERROR] Build failed."
fi

# Cleanup staging folder
rm -rf bin_temp
