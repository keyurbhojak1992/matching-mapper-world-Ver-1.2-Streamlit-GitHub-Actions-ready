Google Colab project folder
===========================

Project path:
/content/drive/MyDrive/Layout

Place these files in that folder:
- layout_party_suggestions_colab.py
- one Layout Stock .xlsx file
- one Sales .xlsx file
- one Actions .xlsx file

Default output:
/content/drive/MyDrive/Layout/Output/Salesperson_Layout_Suggestions.zip

Run in Google Colab
-------------------

Cell 1:
from google.colab import drive
drive.mount('/content/drive')

Cell 2:
!pip -q install pandas openpyxl

Cell 3 - automatic file detection:
!python "/content/drive/MyDrive/Layout/layout_party_suggestions_colab.py" --tolerance 1

Cell 3 alternative - exact filenames:
!python "/content/drive/MyDrive/Layout/layout_party_suggestions_colab.py" \
  --layout "Layout Stock sample file.xlsx" \
  --sales "sales sample file.xlsx" \
  --actions "Actions sample file.xlsx" \
  --tolerance 1

Tolerance:
- 0 = Color, Clarity and Flo must all be exact.
- 1 = Color, Clarity and Flo may each differ by up to one rank.

Recommendation defaults:
- Minimum stones = maximum of 5 or 10 percent of the layout.
- Minimum action weight = 45.

Optional changes:
--min-layout-pct 0.10
--min-absolute-stones 5
--min-action-weight 45
--skip-empty-salespersons
--layout-columns-only

The script automatically creates the Output folder when needed.
