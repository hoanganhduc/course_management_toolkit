# Docker Container for Running course

## Overview

This guide explains how to use the `course` CLI inside a Docker container. The
container includes the required dependencies, so you can manage rosters, grading,
OCR, and LMS workflows without manual setup.

## Quick Start

### 1. Pull and Run the Prebuilt Image

```bash
docker pull ghcr.io/hoanganhduc/course:latest
docker run -it --rm -v $(pwd):/workspace ghcr.io/hoanganhduc/course:latest
```

This mounts your current directory to `/workspace` inside the container so that
`students.db` and exports are written to your host.

### 2. Build and Run Locally

```bash
docker build -t course .
docker run -it --rm -v $(pwd):/workspace course
```

### 3. Run in Detached Mode with Persistent Storage

```bash
docker run -d \
    --name course-container \
    --restart always \
    -v $HOME/Downloads:/home/vscode/Downloads \
    -v $HOME/.config/course:/home/vscode/.config/course \
    -v $(pwd):/workspace \
    ghcr.io/hoanganhduc/course:latest
```

This setup persists downloads and configuration across restarts.

## Running course Commands

```bash
docker exec -it course-container course --help
```

### Optional: Create a Convenience Script

Create a script at `~/.local/bin/course`:

```bash
#!/bin/bash
CONTAINER_NAME="course-container"

if [ $# -lt 1 ]; then
    echo "Usage: $0 [arguments...]"
    exit 1
fi

COMMAND=("course" "$@")

if ! command -v docker &> /dev/null; then
    echo "Error: Docker is not installed"
    exit 1
fi

if ! docker ps -q -f name="$CONTAINER_NAME" | grep -q .; then
    echo "Error: Container '$CONTAINER_NAME' is not running"
    exit 1
fi

docker exec -i "$CONTAINER_NAME" "${COMMAND[@]}"
```

Make it executable:

```bash
chmod +x ~/.local/bin/course
```

Now you can run `course` directly from your terminal:

```bash
course --help
```

---

For more information, see the main repository.
