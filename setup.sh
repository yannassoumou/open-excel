#!/bin/bash
# setup.sh - Unified setup for kuroagent Excel Add-in
#
# Usage:
#   curl -fsSL https://raw.githubusercontent.com/yannassoumou/open-excel/master/setup.sh | bash
#   # OR from local:
#   ./setup.sh              # Interactive menu
#   ./setup.sh install      # First-time install
#   ./setup.sh update       # Pull latest + refresh
#   ./setup.sh uninstall    # Remove CLI + stop server
#   ./setup.sh purge        # Uninstall + delete .kuroagent

set -e

# --- Colors ---
CYAN='\033[0;36m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
WHITE='\033[1;37m'
NC='\033[0m'

ADDIN_NAME="kuroagent"
REPO_URL="https://github.com/yannassoumou/open-excel.git"
INSTALL_DIR="$HOME/.kuroagent"
ACTION="${1:-}"

print_step()    { echo -e "\n${CYAN}>>> $1${NC}"; }
print_success() { echo -e "    ${GREEN}[OK] $1${NC}"; }
print_warning() { echo -e "    ${YELLOW}[!!] $1${NC}"; }
print_error()   { echo -e "    ${RED}[!!] $1${NC}"; }

# --- Interactive menu ---
ask_action() {
    echo -e ""
    echo -e "${CYAN}  ================================${NC}"
    echo -e "${CYAN}  ${ADDIN_NAME} Setup${NC}"
    echo -e "${CYAN}  ================================${NC}"
    echo -e ""
    echo -e "  1) Install   - First-time setup"
    echo -e "  2) Update    - Pull latest and refresh"
    echo -e "  3) Uninstall - Remove CLI + stop server"
    echo -e "  4) Purge     - Uninstall + delete .kuroagent"
    echo -e ""
    read -rp "  Choose [1-4]: " choice
    case "$choice" in
        1) ACTION="install" ;;
        2) ACTION="update" ;;
        3) ACTION="uninstall" ;;
        4) ACTION="purge" ;;
        *) print_error "Invalid choice. Aborting."; exit 1 ;;
    esac
}

# ==================== INSTALL ====================
do_install() {
    print_step "Installing ${ADDIN_NAME}"

    if [ -d "$INSTALL_DIR" ]; then
        print_step "Repository already exists at $INSTALL_DIR"
        print_step "Pulling latest changes ..."
        cd "$INSTALL_DIR" && git pull origin master 2>/dev/null || true
        print_success "Repository updated"
    else
        print_step "Cloning repository ..."
        git clone "$REPO_URL" "$INSTALL_DIR" 2>/dev/null
        print_success "Cloned to $INSTALL_DIR"
        cd "$INSTALL_DIR"
    fi

    # Check Node.js
    print_step "Checking prerequisites"
    if command -v node &>/dev/null; then
        print_success "Node.js $(node -v)"
    else
        print_warning "Node.js not found. Installing via nvm ..."
        export NVM_DIR="$HOME/.nvm"
        [ -s "$NVM_DIR/nvm.sh" ] && . "$NVM_DIR/nvm.sh"
        if command -v nvm &>/dev/null; then
            nvm install 20
        else
            curl -o- https://raw.githubusercontent.com/nvm-sh/nvm/v0.40.1/install.sh | bash
            export NVM_DIR="$HOME/.nvm"
            [ -s "$NVM_DIR/nvm.sh" ] && . "$NVM_DIR/nvm.sh"
            nvm install 20
        fi
        print_success "Node.js $(node -v) via nvm"
    fi

    if command -v npm &>/dev/null; then
        print_success "npm $(npm -v)"
    else
        print_error "npm not available. Aborting."
        exit 1
    fi

    # Install dependencies
    print_step "Installing npm dependencies"
    npm install --no-audit --no-fund --loglevel=error
    print_success "Dependencies installed"

    # Dev certs
    print_step "Installing dev certificates"
    npx office-addin-dev-certs install --machine 2>/dev/null && \
        print_success "Dev certificates installed" || \
        print_warning "Dev certs step returned non-zero (may already be installed)"

    # Register CLI
    print_step "Registering 'kuroagent' CLI"
    if npm link 2>/dev/null; then
        print_success "npm link -- kuroagent CLI on PATH"
    else
        BIN_DIR="$HOME/.local/bin"
        mkdir -p "$BIN_DIR"
        ln -sf "$INSTALL_DIR/bin/kuroagent" "$BIN_DIR/kuroagent"
        if [[ ":$PATH:" != *":$BIN_DIR:"* ]]; then
            SHELL_RC="$HOME/.bashrc"
            [ -f "$HOME/.zshrc" ] && SHELL_RC="$HOME/.zshrc"
            if ! grep -q '.local/bin' "$SHELL_RC" 2>/dev/null; then
                echo 'export PATH="$HOME/.local/bin:$PATH"' >> "$SHELL_RC"
            fi
            export PATH="$HOME/.local/bin:$PATH"
        fi
        print_success "Symlinked kuroagent CLI to $BIN_DIR/kuroagent"
    fi

    # Verify
    print_step "Verifying installation"
    if command -v kuroagent &>/dev/null; then
        print_success "kuroagent CLI found"
        kuroagent --version
    else
        print_error "kuroagent CLI not on PATH. Refresh terminal or restart."
        exit 1
    fi

    # Manifest
    print_step "Checking manifest"
    npx office-addin-manifest validate manifest.xml 2>/dev/null && \
        print_success "manifest.xml valid" || \
        print_warning "Manifest validation returned non-zero (may be OK for dev)"

    show_summary "installed"
}

# ==================== UPDATE ====================
do_update() {
    print_step "Updating ${ADDIN_NAME}"

    if [ ! -d "$INSTALL_DIR" ]; then
        print_error "Not installed. Run install first."
        exit 1
    fi

    cd "$INSTALL_DIR"

    # Pull
    print_step "Pulling latest changes"
    git fetch origin master 2>/dev/null
    LOCAL=$(git rev-parse HEAD)
    REMOTE=$(git rev-parse FETCH_HEAD)
    if [ "$LOCAL" = "$REMOTE" ]; then
        print_success "Already up to date"
    else
        git reset --hard FETCH_HEAD 2>/dev/null
        print_success "Pull complete"
    fi

    # Reinstall deps
    print_step "Reinstalling dependencies"
    npm install --no-audit --no-fund --loglevel=error
    print_success "Dependencies updated"

    # Re-link
    print_step "Re-linking kuroagent CLI"
    npm uninstall -g kuroagent-custom-functions-js 2>/dev/null || true
    npm link 2>/dev/null && print_success "CLI re-linked" || \
        print_warning "npm link returned non-zero (may still work)"

    # Verify
    if command -v kuroagent &>/dev/null; then
        print_step "Verifying"
        print_success "kuroagent CLI: $(kuroagent --version)"
    fi

    show_summary "updated"
}

# ==================== UNINSTALL ====================
do_uninstall() {
    local remove_dir="${1:-false}"

    print_step "Uninstalling ${ADDIN_NAME}"

    # Unlink CLI
    print_step "Unlinking kuroagent CLI"
    if npm uninstall -g kuroagent-custom-functions-js 2>/dev/null; then
        print_success "npm uninstall completed"
    else
        print_warning "CLI was not linked via npm"
    fi

    # Remove manual symlink
    SYMLINK="$HOME/.local/bin/kuroagent"
    if [ -L "$SYMLINK" ]; then
        rm -f "$SYMLINK"
        print_success "Removed symlink: $SYMLINK"
    fi

    # Stop dev server
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

    # Install dir
    if [ "$remove_dir" = "true" ]; then
        print_step "Removing installation directory: $INSTALL_DIR"
        if [ -d "$INSTALL_DIR" ]; then
            rm -rf "$INSTALL_DIR"
            print_success "Directory removed"
        else
            print_warning "Directory not found"
        fi
    else
        print_warning "Keeping $INSTALL_DIR (use 'purge' to remove)"
    fi

    echo -e "\n${GREEN}========================================${NC}"
    echo -e "${GREEN}  ${ADDIN_NAME} uninstalled successfully${NC}"
    echo -e "${GREEN}========================================${NC}"
    echo -e ""
    echo -e "  To reinstall:"
    echo -e "    ${CYAN}curl -fsSL https://raw.githubusercontent.com/yannassoumou/open-excel/master/setup.sh | bash${NC}"
    echo -e ""
}

# ==================== SUMMARY ====================
show_summary() {
    local mode="$1"
    echo -e "\n${GREEN}========================================${NC}"
    echo -e "${GREEN}  ${ADDIN_NAME} ${mode} successfully${NC}"
    echo -e "${GREEN}========================================${NC}"
    echo -e ""
    echo -e "  Quick start:"
    echo -e "    ${CYAN}kuroagent                  Start dev server + sideload${NC}"
    echo -e "    ${CYAN}kuroagent -f file.xlsx     Start + open workbook${NC}"
    echo -e "    ${CYAN}kuroagent --stop           Stop dev server${NC}"
    echo -e "    ${CYAN}kuroagent -m path.xml      Use custom manifest${NC}"
    echo -e "    ${CYAN}kuroagent --no-open        Server only${NC}"
    echo -e ""
    echo -e "  Dev server runs on: ${CYAN}https://localhost:3000${NC}"
    echo -e "  Repo location:      ${CYAN}$INSTALL_DIR${NC}"
    echo -e ""
}

# ==================== HELP ====================
show_help() {
    echo -e ""
    echo -e "  ${ADDIN_NAME} Unified Setup"
    echo -e ""
    echo -e "  Usage:"
    echo -e "    setup.sh              Interactive menu"
    echo -e "    setup.sh install      First-time installation"
    echo -e "    setup.sh update       Pull latest + refresh deps"
    echo -e "    setup.sh uninstall    Remove CLI + stop server"
    echo -e "    setup.sh purge        Uninstall + delete .kuroagent"
    echo -e ""
}

# ==================== DISPATCH ====================
[ -z "$ACTION" ] && ask_action

case "$ACTION" in
    install)   do_install ;;
    update)    do_update ;;
    uninstall) do_uninstall false ;;
    purge)     do_uninstall true ;;
    help)      show_help ;;
    *)         print_error "Unknown action: $ACTION"; show_help; exit 1 ;;
esac
