# GEP Weather to Video

A Python automation tool that captures weather forecasts and converts them into PowerPoint presentations and videos.

## Overview

This tool automates the process of:
1. Capturing weather forecast data from freemeteo.gr
2. Inserting the data into a PowerPoint presentation
3. Converting the presentation into a video file

## Requirements

- Python 3.12+
- Google Chrome browser
- Microsoft PowerPoint

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/powerpoint_automator.git
cd powerpoint_automator
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the executable or script:
```bash
python main.py
```

2. The tool will:
    - Open Chrome and navigate to the weather forecast
    - Take a screenshot of the forecast data
    - Update the PowerPoint template with the new data
    - Convert the presentation to video
    - Save the video to your Desktop as "Καιρός-ΜΗΝΑΣ-2024.mp4"

## Configuration

The default configuration can be found in `main.py`:

```python
CONFIG = {
     'url': "https://freemeteo.gr/kairos/plati/...",
     'screenshot_path': "weather_screenshot.png",
     'template_path': "template.pptx",
     'output_pptx': "updated_presentation.pptx",
     'output_video': "Καιρός-ΜΗΝΑΣ-2024.mp4",
}
```

## Development

### Running Tests

```bash
python -m pytest
```

### Building the Executable

```bash
pyinstaller --onefile --icon=./assets/weather-news.ico --name "GEP Weather to Video" --noconsole --windowed main.py
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Support

If you encounter any problems, please file an issue on the GitHub repository or contact me via email.
