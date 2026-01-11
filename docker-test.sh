#!/bin/bash
# Test Docker build locally

echo "🐳 Testing Docker build..."
docker build -t survey-app:test .

if [ $? -eq 0 ]; then
    echo "✅ Build successful!"
    echo ""
    echo "To test run:"
    echo "  docker run -d -p 8000:8000 -v \$(pwd)/data:/app/data --name survey-test survey-app:test"
    echo "  docker logs -f survey-test"
    echo ""
    echo "To stop:"
    echo "  docker stop survey-test && docker rm survey-test"
else
    echo "❌ Build failed"
    exit 1
fi
