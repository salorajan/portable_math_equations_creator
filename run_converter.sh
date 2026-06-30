#!/bin/bash
# ===================================================
#             EQUATION CONVERTER UTILITY (macOS)
# ===================================================

echo "==================================================="
echo "            EQUATION CONVERTER UTILITY"
echo "==================================================="
echo ""

INPUT_FILE="$1"

if [ -z "$INPUT_FILE" ]; then
    read -p "Enter the path of the file to convert: " INPUT_FILE
fi

# Clean quotes from path
INPUT_FILE=$(echo "$INPUT_FILE" | tr -d '"'\')

if [ ! -f "$INPUT_FILE" ]; then
    echo "[ERROR] Input file \"$INPUT_FILE\" does not exist."
    exit 1
fi

echo ""
echo "Input file selected: $INPUT_FILE"
echo ""
echo "Select the target output format:"
echo " [1] MS Word (.docx)"
echo " [2] HTML (.html)"
echo " [3] PDF (.pdf)"
echo " [4] Markdown (.md)"
echo ""
read -p "Enter choice (1-4): " FORMAT_CHOICE

case "$FORMAT_CHOICE" in
    1) FORMAT="word" ;;
    2) FORMAT="html" ;;
    3) FORMAT="pdf" ;;
    4) FORMAT="md" ;;
    *) echo "[ERROR] Invalid choice." ; exit 1 ;;
esac

# Resolve runner based on OS and architecture
OS="$(uname -s)"
ARCH="$(uname -m)"

if [ "$OS" = "Darwin" ]; then
    if [ "$ARCH" = "arm64" ] && [ -f "bin/convert_md_mac_arm64" ]; then
        RUNNER="bin/convert_md_mac_arm64"
    elif [ -f "bin/convert_md_mac_x86_64" ]; then
        RUNNER="bin/convert_md_mac_x86_64"
    elif [ -f "bin/convert_md_mac" ]; then
        RUNNER="bin/convert_md_mac"
    elif [ -f "references/convert_md.py" ]; then
        RUNNER="python3 references/convert_md.py"
    else
        RUNNER="python3 convert_md.py"
    fi
elif [ "$OS" = "Linux" ]; then
    if [ -f "bin/convert_md_linux" ]; then
        RUNNER="bin/convert_md_linux"
    elif [ -f "references/convert_md.py" ]; then
        RUNNER="python3 references/convert_md.py"
    else
        RUNNER="python3 convert_md.py"
    fi
else
    if [ -f "references/convert_md.py" ]; then
        RUNNER="python3 references/convert_md.py"
    else
        RUNNER="python3 convert_md.py"
    fi
fi

echo ""
echo "Running conversion: $RUNNER \"$INPUT_FILE\" $FORMAT"
echo ""

$RUNNER "$INPUT_FILE" $FORMAT

if [ $? -eq 0 ]; then
    echo ""
    echo "[SUCCESS] Conversion completed successfully!"
else
    echo ""
    echo "[ERROR] Conversion failed."
fi
