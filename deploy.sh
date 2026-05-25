#!/bin/bash

if [[ "${1}" == "-h" || "${1}" == "--help" ]]; then
    echo ""
    echo "Usage: ./deploy.sh --host HOST --user USER --password-hash HASH --secret-key KEY [OPTIONS]"
    echo ""
    echo "  Deploys Nash Analysis to a remote Linux server using Docker via Ansible."
    echo "  Clones the repo on the server, writes .env, and runs docker compose up --build."
    echo ""
    echo "Required:"
    echo "  --host HOST             Server IP or hostname"
    echo "  --user USER             SSH username on the remote server"
    echo "  --password-hash HASH    bcrypt password hash (b64:... from make_password.py)"
    echo "  --secret-key KEY        Flask secret key (generate with:"
    echo "                          python3 -c \"import secrets; print(secrets.token_urlsafe(48))\")"
    echo ""
    echo "Options:"
    echo "  -h, --help              Show this help message"
    echo "  --port PORT             External port the app listens on (default: 5051)"
    echo "  --username USER         App login username (default: admin)"
    echo "  --branch BRANCH         Git branch to deploy (default: main)"
    echo "  --app-dir DIR           Install directory on server (default: /opt/nash-events)"
    echo "  --ssh-key PATH          Path to SSH private key (default: ~/.ssh/id_rsa)"
    echo "  --ssh-port PORT         SSH port on the remote server (default: 22)"
    echo "  --sudo-pass PASS        Sudo password on the remote server"
    echo "  --passwordless-sudo     Server has passwordless sudo — skip sudo prompt"
    echo "  --redeploy              Pull latest code and restart only (skip Docker install)"
    echo "  --check                 Dry run — show what would change without applying"
    echo "  --verbose               Verbose Ansible output (-v)"
    echo ""
    echo "Examples:"
    echo "  ./deploy.sh --host 192.168.1.100 --user ubuntu \\"
    echo "              --password-hash b64:... --secret-key abc123..."
    echo ""
    echo "  ./deploy.sh --host 192.168.1.100 --user ubuntu --port 80 \\"
    echo "              --password-hash b64:... --secret-key abc123... \\"
    echo "              --passwordless-sudo"
    echo ""
    echo "  # Re-deploy only (skip Docker install):"
    echo "  ./deploy.sh --host 192.168.1.100 --user ubuntu \\"
    echo "              --password-hash b64:... --secret-key abc123... \\"
    echo "              --redeploy"
    echo ""
    echo "Generate credentials:"
    echo "  python make_password.py                                      # password hash"
    echo "  python3 -c \"import secrets; print(secrets.token_urlsafe(48))\" # secret key"
    echo ""
    exit 0
fi

cd "$(dirname "$0")"

if ! command -v ansible-playbook &>/dev/null; then
    echo "ERROR: ansible-playbook not found."
    echo "Install: pip install ansible"
    exit 1
fi

# Parse arguments
HOST=""
USER=""
PASSWORD_HASH=""
SECRET_KEY=""
PORT=""
USERNAME=""
BRANCH=""
APP_DIR=""
SSH_KEY=""
SSH_PORT=""
SUDO_PASS=""
PASSWORDLESS_SUDO=""
REDEPLOY=""
CHECK=""
VERBOSE=""
ANSIBLE_ARGS=()

while [[ $# -gt 0 ]]; do
    case "${1}" in
        --host)              HOST="${2}";           shift 2 ;;
        --user)              USER="${2}";           shift 2 ;;
        --password-hash)     PASSWORD_HASH="${2}";  shift 2 ;;
        --secret-key)        SECRET_KEY="${2}";     shift 2 ;;
        --port)              PORT="${2}";           shift 2 ;;
        --username)          USERNAME="${2}";       shift 2 ;;
        --branch)            BRANCH="${2}";         shift 2 ;;
        --app-dir)           APP_DIR="${2}";        shift 2 ;;
        --ssh-key)           SSH_KEY="${2}";        shift 2 ;;
        --ssh-port)          SSH_PORT="${2}";       shift 2 ;;
        --sudo-pass)         SUDO_PASS="${2}";      shift 2 ;;
        --passwordless-sudo) PASSWORDLESS_SUDO=1;   shift ;;
        --redeploy)          REDEPLOY=1;            shift ;;
        --check)             CHECK=1; ANSIBLE_ARGS+=("--check" "--diff"); shift ;;
        --verbose)           VERBOSE=1; ANSIBLE_ARGS+=("-v"); shift ;;
        *)                   ANSIBLE_ARGS+=("${1}"); shift ;;
    esac
done

# Validate required args
if [[ -z "$HOST" || -z "$USER" || -z "$PASSWORD_HASH" || -z "$SECRET_KEY" ]]; then
    echo "ERROR: --host, --user, --password-hash, and --secret-key are required."
    echo "Usage: ./deploy.sh --host HOST --user USER --password-hash HASH --secret-key KEY"
    echo "       ./deploy.sh --help   for full usage"
    exit 1
fi

# Write a temporary inventory file
TMPINV="$(mktemp /tmp/nash-inventory.XXXXXX)"
trap "rm -f ${TMPINV}" EXIT

{
    echo "[nash_events]"
    printf '%s ansible_user=%s' "${HOST}" "${USER}"
    [[ -n "$SSH_KEY"  ]] && printf ' ansible_ssh_private_key_file=%s' "${SSH_KEY}"
    [[ -n "$SSH_PORT" ]] && printf ' ansible_port=%s' "${SSH_PORT}"
    echo ""
} > "${TMPINV}"

# Sudo handling
if [[ -n "$SUDO_PASS" ]]; then
    ANSIBLE_ARGS+=("-e" "ansible_become_password=${SUDO_PASS}")
elif [[ -z "$PASSWORDLESS_SUDO" ]]; then
    ANSIBLE_ARGS+=("--ask-become-pass")
fi

# Build extra vars
ANSIBLE_ARGS+=("-e" "nash_password_hash=${PASSWORD_HASH}")
ANSIBLE_ARGS+=("-e" "secret_key=${SECRET_KEY}")
[[ -n "$PORT"     ]] && ANSIBLE_ARGS+=("-e" "app_port=${PORT}")
[[ -n "$USERNAME" ]] && ANSIBLE_ARGS+=("-e" "nash_username=${USERNAME}")
[[ -n "$BRANCH"   ]] && ANSIBLE_ARGS+=("-e" "app_branch=${BRANCH}")
[[ -n "$APP_DIR"  ]] && ANSIBLE_ARGS+=("-e" "app_dir=${APP_DIR}")
[[ -n "$REDEPLOY" ]] && ANSIBLE_ARGS+=("--tags" "deploy")

# Summary
[[ -n "$CHECK" ]] && echo "╔══════════════════════════════════════╗" && \
                     echo "║         DRY RUN — no changes made    ║" && \
                     echo "╚══════════════════════════════════════╝"
echo "Deploying Nash Analysis..."
echo "  Host    : ${HOST}"
echo "  User    : ${USER}"
[[ -n "$SSH_PORT"  ]] && echo "  SSH port: ${SSH_PORT}"
echo "  Port    : ${PORT:-5051}"
echo "  Branch  : ${BRANCH:-main}"
echo "  Dir     : ${APP_DIR:-/opt/nash-events}"
[[ -n "$PASSWORDLESS_SUDO" ]] && echo "  Sudo    : passwordless"
[[ -n "$SUDO_PASS"         ]] && echo "  Sudo    : password supplied"
[[ -z "$SUDO_PASS" && -z "$PASSWORDLESS_SUDO" ]] && echo "  Sudo    : will prompt"
[[ -n "$REDEPLOY"  ]] && echo "  Mode    : redeploy only (skip Docker install)"
[[ -n "$CHECK"     ]] && echo "  Mode    : DRY RUN (--check --diff)"
[[ -n "$VERBOSE"   ]] && echo "  Verbose : yes"
echo ""

ansible-playbook deploy/ansible/deploy_docker.yml -i "${TMPINV}" "${ANSIBLE_ARGS[@]}"
