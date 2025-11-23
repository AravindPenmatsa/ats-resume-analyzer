#!/bin/bash

# Colors for output
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo -e "${GREEN}Starting ATS Resume Analyzer Setup...${NC}"

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "Python 3 is not installed. Please install it first."
    exit 1
fi

# Create virtual environment if it doesn't exist
if [ ! -d "venv" ]; then
    echo -e "${YELLOW}Creating virtual environment...${NC}"
    python3 -m venv venv
fi

# Activate virtual environment
echo -e "${YELLOW}Activating virtual environment...${NC}"
source venv/bin/activate

# Install dependencies
echo -e "${YELLOW}Installing dependencies...${NC}"
pip install -r requirements.txt

# Check for .env file
if [ ! -f ".env" ]; then
    echo -e "${YELLOW}Warning: .env file not found!${NC}"
    echo "Please create a .env file with your OPENAI_API_KEY."
    echo "Example: OPENAI_API_KEY=your_key_here"
    # Optional: Create a dummy .env if needed, or just warn
    # touch .env
fi

# Run the application
echo -e "${GREEN}Starting the application...${NC}"
echo "Access the app at http://localhost:8000"
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
