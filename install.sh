#!/bin/bash
# install.sh — Install kuroagent Excel Add-in for development
#
# Installs dependencies, registers the `kuroagent` CLI on PATH, and
# verifies the dev server can start.
#
# Usage:
#   curl -sSL https://raw.githubusercontent.com/yannassoumou/kuroagent/main/install.sh | bash
#   # OR from local repo:
#   ./install.sh

set -e

# --- Colors ---
CYAN='\033[0;36m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
WHITE='\033[1;37m'
NC='\033[0m'

ADDIN_NAME="kuroagent"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

print_step()    { echo -e "\n${CYAN}>>> $1${NC}"; }
print_success() { echo -e "    ${GREEN}[OK] $1${NC}"; }
print_warning() { echo -e "    ${YELLOW}[!!] $1${NC}"; }
print_error()   { echo -e "    ${RED}[!!] $1${NC}"; }

# --- Detect mode ---
# If run from the repo, use local dir.  If piped via curl, clone the repo.
if [ -f "$SCRIPT_DIR/package.json" ]; then
    REPO_DIR="$SCRIPT_DIR"
    print_step "Running in existing repo: $REPO_DIR"
else
    CLONE_DIR="$HOME/.kuroagent"
    if [ ! -d "$CLONE_DIR" ]; then
        print_step "Cloning repository ..."
        git clone https://github.com/yannassoumou/open-excel.git "$CLONE_DIR"
        print_success "Cloned to $CLONE_DIR"
    else
        print_step "Repository already cloned, pulling latest ..."
        cd "$CLONE_DIR" && git pull 2>/dev/null || true
    fi
    REPO_DIR="$CLONE_DIR"
fi

cd "$REPO_DIR"

# --- Check Node.js ---
print_step "Checking prerequisites"

if command -v node &>/dev/null; then
    NODE_VER=$(node -v)
    print_success "Node.js $NODE_VER"
else
    print_warning "Node.js not found. Installing via nvm ..."
    curl -o- https://raw.githubusercontent.com/nvm-sh/nvm/v0.40.1/install.sh | bash
    export NVM_DIR="$HOME/.nvm"
    [ -s "$NVM_DIR/nvm.sh" ] && . "$NVM_DIR/nvm.sh"
    nvm install 20
    print_success "Node.js $(node -v) via nvm"
fi

if command -v npm &>/dev/null; then
    print_success "npm $(npm -v)"
else
    print_error "npm not available. Aborting."
    exit 1
fi

# --- Install dependencies ---
print_step "Installing dependencies"
npm install --no-audit --no-fund --loglevel=error
print_success "Dependencies installed"

# --- Install dev certs (required for HTTPS webpack server) ---
print_step "Installing dev certificates"
if command -v npx &>/dev/null; then
    npx office-addin-dev-certs install --machine 2>/dev/null && \
        print_success "Dev certificates installed" || \
        print_warning "Dev certs step returned non-zero (may already be installed)"
fi

# --- Register the `kuroagent` CLI ---
print_step "Registering 'kuroagent' CLI"

# Option 1: npm link (works when cwd is the repo)
if npm link 2>/dev/null; then
    print_success "npm link -- kuroagent CLI on PATH"
else
    # Option 2: manual symlink
    BIN_DIR="$HOME/.local/bin"
    mkdir -p "$BIN_DIR"
    ln -sf "$REPO_DIR/bin/kuroagent" "$BIN_DIR/kuroagent"
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

# --- Verify ---
print_step "Verifying installation"
if command -v kuroagent &>/dev/null; then
    print_success "kuroagent CLI found: $(which kuroagent)"
    kuroagent --version
else
    print_error "kuroagent CLI not on PATH. Add $(npm bin -g 2>/dev/null || echo global npm dir) to PATH."
    exit 1
fi

# --- Manifest check ---
print_step "Checking manifest"
if command -v npx &>/dev/null; then
    npx office-addin-manifest validate manifest.xml 2>/dev/null && \
        print_success "manifest.xml valid" || \
        print_warning "Manifest validation returned non-zero (may be OK for dev)"
fi

# --- Summary ---
echo -e "\n${GREEN}════════════════════════════════════════${NC}"
echo -e "${GREEN}  ${ADDIN_NAME} installed successfully${NC}"
echo -e "${GREEN}════════════════════════════════════════${NC}"
echo -e ""
echo -e "  Quick start:"
echo -e "    ${CYAN}kuroagent${NC}                  Start dev server + sideload"
echo -e "    ${CYAN}kuroagent --stop${NC}           Stop dev server"
echo -e "    ${CYAN}kuroagent -m path/to.xml${NC}   Use custom manifest"
echo -e "    ${CYAN}kuroagent --no-open${NC}        Server only"
echo -e ""
echo -e "  Dev server runs on: ${CYAN}https://localhost:3000${NC}"
echo -e "  Repo location:      ${CYAN}$REPO_DIR${NC}"
echo -e ""
