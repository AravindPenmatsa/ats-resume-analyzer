# ATS Resume Analyzer

A tool to evaluate and optimize resumes for ATS compatibility.

## Getting Started

### Prerequisites
- Python 3.8 or higher
- pip (Python package manager)

### Quick Start (Mac/Linux)
1.  **Clone the repository** (if you haven't already):
    ```bash
    git clone <repository-url>
    cd ats-resume-analyzer
    ```

2.  **Run the setup script**:
    ```bash
    ./run_app.sh
    ```
    This script will automatically:
    - Create a virtual environment
    - Install dependencies
    - Start the application

### Manual Setup
If you prefer to run commands manually:

1.  **Create a virtual environment**:
    ```bash
    python3 -m venv venv
    source venv/bin/activate
    ```

2.  **Install dependencies**:
    ```bash
    pip install -r requirements.txt
    ```

3.  **Environment Configuration**:
    - Create a `.env` file in the root directory.
    - Add your OpenAI API key:
      ```
      OPENAI_API_KEY=your_api_key_here
      ```

4.  **Run the application**:
    ```bash
    uvicorn app.main:app --reload
    ```

### Access the App
Open your browser and navigate to: [http://localhost:8000](http://localhost:8000)