#!/bin/bash
# Build script for Docker container

echo "🐳 Building Docker image..."
sudo docker build -t survey-app:latest .

echo "✅ Build complete!"
echo ""
echo "To run:"
echo "  docker run -d -p 8000:8000 -v $(pwd)/data:/app/data --name survey-app survey-app:latest"
echo ""
echo "Or use docker-compose:"
echo "  docker-compose up -d"

