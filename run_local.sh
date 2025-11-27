#!/bin/bash
# Local development server

echo "Starting local Azure Functions host..."
echo "Bot will be available at: http://localhost:7071/api/messages"
echo ""
echo "Press Ctrl+C to stop"
echo ""

# Activate virtual environment if it exists
if [ -d "venv" ]; then
    source venv/bin/activate
fi

# Start the function host
func start --python
