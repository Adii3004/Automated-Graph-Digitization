# Automated Graph Digitization and Performance Analysis
A Python-based tool for extracting numerical data from performance graphs (pumps & compressors), fitting curves and generating performance reports._

---

## Project overview
This repository contains the implementation of an interactive application that:
- Loads graph images (PNG, JPG, BMP),
- Lets users manually trace curves (click points) on the displayed graph,
- Transforms pixel coordinates to real-world axis values based on user-supplied axis limits,
- Performs curve fitting (polynomial or spline) and interpolation,
- Calculates performance metrics (e.g., flow, head, efficiency),
- Exports results (raw & transformed points, fitted curve data, plots) to an Excel report.

Use cases: pump/compressor performance analysis, digitizing plot data from papers, lab reports and legacy documents.

---

## Quick demo
```bash
# create environment
python -m venv venv
source venv/bin/activate   # On Windows: venv\Scripts\activate

# install dependencies
pip install -r requirements.txt

# run GUI app
python src/app.py
```

---

## Repo structure
```
/
├─ src/
│  └─ curvetracing.py          # Main Python script containing GUI, tracing, and analysis logic
│
├─ examples/
│  ├─ demo_graph.png           # Sample input graph image
│  ├─ output_trace_1.xlsx      # Example Excel output file
│  ├─ output_trace_2.xlsx      # Another example Excel output
│  ├─ performance_data.xlsx    # Example performance analysis data
│
│
├─ requirements.txt            # Dependencies list
├─ README.md                   # Project documentation

```

---

## Features
- Image zooming & panning for precise point selection.
- Real-time visual feedback for selected points.
- Undo / redo last point.
- Two fitting modes: polynomial (configurable degree) and cubic spline.
- Export: `.xlsx` containing raw pixel points, transformed values, fitted curve sampled points, and plots embedded in the sheet.

---

## Installation (detailed)
1. Clone the repo:
```bash
git clone https://github.com/<your-username>/automated-graph-digitization.git
cd automated-graph-digitization
```

2. Create & activate a virtual environment:
```bash
python3 -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

---

## Usage (GUI)
1. `python src/app.py`  
2. In the GUI:
   - `Load Image` → select a graph image.
   - `Set Axes` → input X_min, X_max, Y_min, Y_max (real-world axis values).
   - `Start Tracing` → click points along the curve. Use `Undo` to remove last point.
   - `Fit Curve` → choose polynomial degree or spline.
   - `Export Report` → save `.xlsx` containing the results and plots.

---

## Usage (CLI / script example)
You can also run non-GUI pipeline to transform a pre-collected CSV of pixel points and produce a report:
```bash
python src/transform_and_report.py --points examples/demo_points.csv \
  --image examples/demo_graph.png \
  --x-min 0 --x-max 100 --y-min 0 --y-max 10 \
  --fit polynomial --degree 3 \
  --output demo_out.xlsx
```

---

## Tests
Run unit tests with pytest:
```bash
pytest -q
```

---

## Contributing
- Fork the repository.
- Create a feature branch: `git checkout -b feature/my-feature`
- Commit your changes: `git commit -m "Add new feature"`
- Push: `git push origin feature/my-feature`
- Create a Pull Request detailing the changes and test steps.

---

## TODO / Future work
- Automatic curve detection using CNN / image segmentation (ChartOCR-like models).
- Support for logarithmic and polar axes.
- Web UI + cloud processing for batch jobs.
- Digital twin integration for real-time monitoring.


