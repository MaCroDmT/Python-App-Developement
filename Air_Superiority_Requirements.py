import os
from pathlib import Path
from fpdf import FPDF

# --- 1. SETUP AUTOMATIC PATH TO DOWNLOADS ---
def get_downloads_path():
    """Returns the absolute path to the user's Downloads folder."""
    return str(Path.home() / "Downloads")

# Define output filename and path
output_filename = "Air_Superiority_Requirements.pdf"
output_path = os.path.join(get_downloads_path(), output_filename)

# --- 2. PDF CLASS DEFINITION ---
class PDF(FPDF):
    def header(self):
        # Header: Title
        self.set_font('Arial', 'B', 16)
        self.cell(0, 10, 'Project Requirements: Air Superiority (Hangar Manager)', 0, 1, 'C')
        self.ln(5)

    def footer(self):
        # Footer: Page number
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, 'Page ' + str(self.page_no()) + '/{nb}', 0, 0, 'C')

    def section_title(self, title):
        # Section Heading with background color
        self.set_font('Arial', 'B', 12)
        self.set_fill_color(220, 220, 220)  # Light gray
        self.cell(0, 8, title, 0, 1, 'L', 1)
        self.ln(2)

    def section_body(self, txt):
        # Regular text
        self.set_font('Arial', '', 11)
        self.multi_cell(0, 5, txt)
        self.ln()

    def bullet_point(self, txt):
        # Indented bullet point
        self.set_font('Arial', '', 11)
        self.cell(5) # Indent 5mm
        self.cell(5, 5, '-', 0, 0)
        self.multi_cell(0, 5, txt)

# --- 3. GENERATE CONTENT ---
pdf = PDF()
pdf.alias_nb_pages()
pdf.add_page()

# 1. Objective
pdf.section_title("1. Objective")
# Content derived from [cite: 3, 4, 5]
pdf.section_body("Create a multi-page Fighter Jet encyclopedic web application using only pure HTML. The project should feel like a military briefing dossier or a flight hangar menu system with navigation, aircraft profiles, tactical views, and squadron details.")

# Constraints
pdf.section_title("2. Constraints & Technologies")
# [cite_start]Content derived from [cite: 6, 7, 61]
pdf.bullet_point("Allowed Technologies: HTML only (No CSS, No JavaScript).")
pdf.bullet_point("Layout: Must use HTML Tables for all structure.")
pdf.bullet_point("Styling: Must use inline HTML attributes (e.g., width, bgcolor, align).")
pdf.ln(5)

# 3. Features Required
pdf.section_title("3. Features Required")

# Home Page
pdf.set_font('Arial', 'B', 11)
pdf.cell(0, 6, "3.1 Home Page (Main Hangar)", 0, 1)
# [cite_start]Content derived from [cite: 9, 10, 11, 12, 13, 14, 15]
pdf.bullet_point("Title & Icon: 'Air Superiority' with a jet fighter icon.")
pdf.bullet_point("Welcome Message: 'Welcome to the Main Hangar, Pilot'.")
pdf.bullet_point("Audio: Autoplay looped background music (jet ambience).")
pdf.bullet_point("Banner: Large banner image of a fighter jet.")
pdf.bullet_point("Navigation: Table-based menu where each cell is fully clickable.")
pdf.ln(3)

# Navigation Bar
pdf.set_font('Arial', 'B', 11)
pdf.cell(0, 6, "3.2 Navigation Bar (Global)", 0, 1)
# [cite_start]Content derived from [cite: 23, 24, 25]
pdf.bullet_point("Must appear on every single page.")
pdf.bullet_point("Structure: Table row containing Home | Back | Current Page Title.")
pdf.ln(3)

# Pilot Profile
pdf.set_font('Arial', 'B', 11)
pdf.cell(0, 6, "3.3 Pilot Profile (User Info)", 0, 1)
# [cite_start]Content derived from [cite: 26, 27, 28, 29]
pdf.bullet_point("Profile Picture: User/Pilot Avatar.")
pdf.bullet_point("Details: Name, Rank, Country, Favorite Jet.")
pdf.bullet_point("Layout: Table-based info card.")
pdf.ln(3)

# Squadron Roster
pdf.set_font('Arial', 'B', 11)
pdf.cell(0, 6, "3.4 Squadron Roster (Team Info)", 0, 1)
# [cite_start]Content derived from [cite: 30, 31, 32, 33]
pdf.bullet_point("Header: Squadron Emblem and Name.")
pdf.bullet_point("Roster Table: Columns for Jet Image and Model Name.")
pdf.bullet_point("Interaction: Each jet must be clickable to view specs.")
pdf.ln(3)

# Tactical Airspace
pdf.set_font('Arial', 'B', 11)
pdf.cell(0, 6, "3.5 Tactical Airspace (Full Field View)", 0, 1)
# [cite_start]Content derived from [cite: 34, 35, 36, 37]
pdf.bullet_point("Layout: Large HTML table representing the sky/map.")
pdf.bullet_point("Left Side: Friendly Squadron jets.")
pdf.bullet_point("Right Side: Hostile/Opponent jets.")
pdf.bullet_point("Info: Mission Status or Score.")
pdf.ln(3)

pdf.add_page() # -- Page Break --

# Air Marshal
pdf.set_font('Arial', 'B', 11)
pdf.cell(0, 6, "3.6 Air Marshal Info (Manager)", 0, 1)
# [cite_start]Content derived from [cite: 38, 39, 40, 41]
pdf.bullet_point("Profile Card: Commander details.")
pdf.bullet_point("Stats Table: Service Record and History.")
pdf.ln(3)

# Gallery
pdf.set_font('Arial', 'B', 11)
pdf.cell(0, 6, "3.7 Spotter's Gallery", 0, 1)
# [cite_start]Content derived from [cite: 42, 43, 44, 45, 46, 47]
pdf.bullet_point("Layout: Grid layout using tables.")
pdf.bullet_point("Content: High-res images OR muted autoplay video clips.")
pdf.bullet_point("Caption: Jet name or maneuver info.")
pdf.ln(3)

# Formations
pdf.set_font('Arial', 'B', 11)
pdf.cell(0, 6, "3.8 Flight Formations", 0, 1)
# [cite_start]Content derived from [cite: 48, 49, 50]
pdf.bullet_point("List of tactical formations with descriptions.")
pdf.bullet_point("Note: Placeholder for future 3D previews.")
pdf.ln(3)

# Subpages
pdf.set_font('Arial', 'B', 11)
pdf.cell(0, 6, "3.9 Aircraft Spec Subpage", 0, 1)
# [cite_start]Content derived from [cite: 51, 52, 53, 54]
pdf.bullet_point("One unique HTML file per jet model.")
pdf.bullet_point("Info Card: Player/Jet info.")
pdf.bullet_point("Stats Table: Speed, Range, Armament.")
pdf.bullet_point("Back Navigation: Returns to Roster or Airspace depending on team.")
pdf.ln(3)

# 4. Folder Structure
pdf.section_title("4. File & Folder Structure")
# [cite_start]Content derived from [cite: 63, 64, 65, 66]
structure_text = """fighter-jet-hangar/
    index.html          (Home Hub)
    squadron.html       (Team/Roster)
    airspace.html       (Full Field/Map)
    commander.html      (Manager/Marshal)
    gallery.html        (Gallery)
    formations.html     (Formations)
    pilot.html          (User Info)
    assets/
        images/
        audio/
        video/
    jets/               (Individual Profiles)
        f16.html
        su35.html
        ..."""
pdf.set_font('Courier', '', 10)
pdf.multi_cell(0, 5, structure_text)

# --- 4. SAVE FILE ---
try:
    pdf.output(output_path)
    print(f"SUCCESS: PDF generated at: {output_path}")
except PermissionError:
    print(f"ERROR: Could not write to {output_path}. Please close the file if it is open.")
except Exception as e:
    print(f"An error occurred: {e}")