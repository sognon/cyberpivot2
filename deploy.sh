#!/usr/bin/env bash
set -euo pipefail

# Load .env if present
if [ -f ".env" ]; then
  # shellcheck disable=SC1091
  source .env
fi

APP_NAME="${APP_NAME:-cyberpivot}"
ENTRYPOINT="${ENTRYPOINT:-app_cyberpivot_risk.py}"
PORT="${PORT:-8501}"
DOCKER_IMAGE="${DOCKER_IMAGE:-cyberpivot}"
DOCKER_TAG="${DOCKER_TAG:-latest}"
IMAGE_REF="${DOCKER_IMAGE}:${DOCKER_TAG}"

REGISTRY="${REGISTRY:-}"
REGISTRY_USER="${REGISTRY_USER:-}"

REMOTE_USER="${REMOTE_USER:-}"
REMOTE_HOST="${REMOTE_HOST:-}"
REMOTE_PATH="${REMOTE_PATH:-/opt/${APP_NAME}}"

usage() {
  cat <<EOF
CyberPivot™ deploy script

Usage:
  $(basename "$0") build            Build local Docker image
  $(basename "$0") run              Run locally on http://localhost:${PORT}
  $(basename "$0") compose          Up with docker-compose (local)
  $(basename "$0") push             Push image to registry (requires REGISTRY/REGISTRY_USER)
  $(basename "$0") remote           Deploy to remote host via SSH (uses docker-compose)
  $(basename "$0") systemd          Output a sample systemd unit for docker-compose

Environment (use .env file or export):
  APP_NAME, ENTRYPOINT, PORT
  DOCKER_IMAGE, DOCKER_TAG
  REGISTRY, REGISTRY_USER
  REMOTE_USER, REMOTE_HOST, REMOTE_PATH
EOF
}

require() {
  if ! command -v "$1" >/dev/null 2>&1; then
    echo "Missing dependency: $1" >&2
    exit 1
  fi
}

cmd_build() {
  require docker
  echo "==> Building ${IMAGE_REF}"
  docker build -t "${IMAGE_REF}" .
}

cmd_run() {
  require docker
  echo "==> Running ${IMAGE_REF} on port ${PORT}"
  docker run --rm -p "${PORT}:8501" \
    -e STREAMLIT_BROWSER_GATHER_USAGE_STATS=false \
    -v "$(pwd)/data:/app/data" \
    -v "$(pwd)/auth_config.yaml:/app/auth_config.yaml:ro" \
    -v "$(pwd)/users_demo.yaml:/app/users_demo.yaml:ro" \
    --name "${APP_NAME}" "${IMAGE_REF}" \
    streamlit run "${ENTRYPOINT}" --server.address=0.0.0.0 --server.port=8501
}

cmd_compose() {
  require docker
  if command -v docker-compose >/dev/null 2>&1; then
    DC="docker-compose"
  else
    require docker
    DC="docker compose"
  fi
  echo "==> Starting with ${DC}"
  ${DC} up -d
  ${DC} ps
  echo "Open: http://localhost:${PORT}"
}

cmd_push() {
  require docker
  if [ -n "${REGISTRY}" ]; then
    IMAGE_WITH_REG="${REGISTRY}/${IMAGE_REF}"
    echo "==> Tagging ${IMAGE_REF} as ${IMAGE_WITH_REG}"
    docker tag "${IMAGE_REF}" "${IMAGE_WITH_REG}"
    echo "==> Logging in to ${REGISTRY}"
    docker login "${REGISTRY}" -u "${REGISTRY_USER}"
    echo "==> Pushing ${IMAGE_WITH_REG}"
    docker push "${IMAGE_WITH_REG}"
  else
    echo "No REGISTRY specified. Pushing to Docker Hub namespace of ${REGISTRY_USER:-<unset>} if applicable."
    docker push "${IMAGE_REF}"
  fi
}

cmd_remote() {
  require ssh
  require scp
  if [ -z "${REMOTE_USER}" ] || [ -z "${REMOTE_HOST}" ]; then
    echo "REMOTE_USER and REMOTE_HOST are required" >&2
    exit 1
  fi

  # Choose docker compose command on remote
  REMOTE_DC="docker compose"
  ssh "${REMOTE_USER}@${REMOTE_HOST}" "command -v docker-compose >/dev/null 2>&1 && echo compose_legacy || true" | grep -q compose_legacy && REMOTE_DC="docker-compose"

  echo "==> Creating remote path ${REMOTE_PATH}"
  ssh "${REMOTE_USER}@${REMOTE_HOST}" "sudo mkdir -p '${REMOTE_PATH}' && sudo chown -R \$USER:\$USER '${REMOTE_PATH}'"

  echo "==> Copying compose + env"
  scp docker-compose.yml "${REMOTE_USER}@${REMOTE_HOST}:${REMOTE_PATH}/docker-compose.yml"
  if [ -f ".env" ]; then
    scp .env "${REMOTE_USER}@${REMOTE_HOST}:${REMOTE_PATH}/.env"
  fi

  echo "==> Pulling/Building and starting"
  ssh "${REMOTE_USER}@${REMOTE_HOST}" "cd '${REMOTE_PATH}' && ${REMOTE_DC} pull || ${REMOTE_DC} build && ${REMOTE_DC} up -d && ${REMOTE_DC} ps"
  echo "==> Remote deploy done."
}

cmd_systemd() {
  cat <<'UNIT'
# /etc/systemd/system/cyberpivot.service
[Unit]
Description=CyberPivot via docker-compose
After=network-online.target docker.service
Wants=docker.service

[Service]
Type=oneshot
RemainAfterExit=yes
WorkingDirectory=/opt/cyberpivot
ExecStart=/usr/bin/docker compose up -d
ExecStop=/usr/bin/docker compose down
TimeoutStartSec=0

[Install]
WantedBy=multi-user.target
UNIT
}

case "${1:-}" in
  build)   cmd_build ;;
  run)     cmd_run ;;
  compose) cmd_compose ;;
  push)    cmd_push ;;
  remote)  cmd_remote ;;
  systemd) cmd_systemd ;;
  *)       usage ;;
esac
