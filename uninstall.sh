#!/bin/bash
# uninstall.sh - Uninstall KuroAgent Excel Add-in and development tools
#
# Usage:
#   curl -fsSL https://raw.githubusercontent.com/yannassoumou/open-excel/master/uninstall.sh | bash
#   # OR from local repo:
#   ./uninstall.sh
#   ./uninstall.sh --all     # Also remove the installation directory

set -e

# --- Colors ---
CYAN='\033[0;36m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

ADDIN_NAME="KuroAgent"
ARG="${1:-}"

print_step()    { echo -e "\n${CYAN}>>> $1${NC}"; }
print_success() { echo -e "    ${GREEN}[OK] $1${NC}"; }
print_warning() { echo -e "    ${YELLOW}[!!] $1${NC}"; }

# --- 1. Unlink the Excel CLI ---
print_step "Unlinking excel CLI"

if npm unlink -g excel-custom-functions-js 2>/dev/null; then
    print_success "npm unlink completed"
else
    print_warning "CLI was not linked via npm (may have been symlinked)"
fi

# Remove manual symlink if it exists
SYMLINK="$HOME/.local/bin/excel"
if [ -L "$SYMLINK" ]; then
    rm -f "$SYMLINK"
    print_success "Removed symlink: $SYMLINK"
fi

# --- 2. Remove installation directory ---
INSTALL_DIR="$HOME/.kuroagent"

if [ "$ARG" = "--all" ]; then
    print_step "Removing installation directory: $INSTALL_DIR"
    if [ -d "$INSTALL_DIR" ]; then
        rm -rf "$INSTALL_DIR"
        print_success "Directory removed: $INSTALL_DIR"
    else
        print_warning "Directory not found: $INSTALL_DIR"
    fi
else
    print_warning "Keeping $INSTALL_DIR (use --all to remove)"
fi

# --- 3. Stop any running dev server ---
print_step "Stopping dev server if running"

if command -v lsof &>/dev/null; then
    PIDS=$(lsof -ti :3000 2>/dev/null || true)
    if [ -n "$PIDS" ]; then
        echo "$PIDS" | xargs -r kill -9 2>/dev/null || true
        print_success "Stopped dev server on port 3000"
    else
        print_warning "No dev server running on port 3000"
    fi
else
    print_warning "lsof not available, skipping server check"
fi

# --- Summary ---
echo -e "\n${GREEN}════════════════════════════════════════${NC}"
echo -e "${GREEN}  ${ADDIN_NAME} uninstalled successfully${NC}"
echo -e "${GREEN}════════════════════════════════════════${NC}"
echo -e ""
echo -e "  To reinstall:"
echo -e "    ${CYAN}curl -fsSL https://raw.githubusercontent.com/yannassoumou/open-excel/master/install.sh | bash${NC}"
echo -e ""
echo -e "  Note: Restart Excel to complete the uninstall."
