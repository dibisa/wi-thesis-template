#!/usr/bin/env python3
"""
Word-to-LaTeX Thesis Converter
Watches for changes to Word document and converts to LaTeX chapter files.

Usage:
    python word_to_latex.py          # Run watcher continuously
    python word_to_latex.py --once   # Convert once and exit
"""

import os
import re
import subprocess
import sys
import time
from pathlib import Path
from typing import Optional

# Configuration
WORD_FILE = Path(r"C:\Users\dibis\OneDrive\Desktop\Thesis\Manuscripts\Thesis_draft_chapter1_3.docx")
LATEX_DIR = Path(r"C:\Users\dibis\OneDrive\Desktop\Thesis\LaTex\wi-thesis-template")
CHAPTERS_DIR = LATEX_DIR / "chapters"

# Acronym definitions (short -> full form)
ACRONYMS = {
    "NRL": "Natural Rubber Latex",
    "NR": "Natural Rubber",
    "PLA": "Polylactic Acid",
    "PHA": "Polyhydroxyalkanoates",
    "SBR": "Styrene-Butadiene Rubber",
    "TPE": "Thermoplastic Elastomer",
    "TPU": "Thermoplastic Polyurethane",
    "PU": "Polyurethane",
    "EPDM": "Ethylene Propylene Diene Monomer",
    "SIC": "Strain-Induced Crystallization",
    "DLVO": "Derjaguin-Landau-Verwey-Overbeek",
    "SDS": "Sodium Dodecyl Sulfate",
    "DLS": "Dynamic Light Scattering",
    "PDI": "Polydispersity Index",
    "DRC": "Dry Rubber Content",
    "TSC": "Total Solids Content",
    "AM": "Additive Manufacturing",
    "VPP": "Vat Photopolymerization",
    "SLA": "Stereolithography",
    "DLP": "Digital Light Processing",
    "FDM": "Fused Deposition Modeling",
    "FFF": "Fused Filament Fabrication",
    "DIW": "Direct Ink Writing",
    "SLS": "Selective Laser Sintering",
    "LOM": "Laminated Object Manufacturing",
    "HDDA": "1,6-Hexanediol Diacrylate",
    "TMPTA": "Trimethylolpropane Triacrylate",
    "TPO": "Phenylbis(2,4,6-trimethylbenzoyl)-phosphine Oxide",
    "PRE": "Photoresin Emulsion",
    "DEPR": "Dual Emulsion Photoresin",
    "JMRE": "Jammed Micro-Reinforced Elastomer",
    "HLB": "Hydrophilic-Lipophilic Balance",
    "HIPE": "High Internal Phase Emulsion",
    "NMR": "Nuclear Magnetic Resonance",
    "DOSY": "Diffusion-Ordered Spectroscopy",
    "HSQC": "Heteronuclear Single Quantum Coherence",
    "HMBC": "Heteronuclear Multiple Bond Correlation",
    "COSY": "Correlation Spectroscopy",
    "CPMG": "Carr-Purcell-Meiboom-Gill",
    "FTIR": "Fourier Transform Infrared Spectroscopy",
    "GPC": "Gel Permeation Chromatography",
    "SEM": "Scanning Electron Microscopy",
    "TEM": "Transmission Electron Microscopy",
    "ASTM": "American Society for Testing and Materials",
    "ISO": "International Organization for Standardization",
    "OSHA": "Occupational Safety and Health Administration",
    "DOE": "Design of Experiments",
    "CFD": "Computational Fluid Dynamics",
    "LVR": "Linear Viscoelastic Region",
    "UV": "Ultraviolet",
    "O/W": "Oil-in-Water",
    "W/O": "Water-in-Oil",
    "THF": "Tetrahydrofuran",
    "TMS": "Tetramethylsilane",
    "CoA": "Certificate of Analysis",
}

# Chapter configurations
CHAPTERS = {
    1: {
        "title": "Introduction",
        "label": "ch:introduction",
        "file": "chapter1.tex",
    },
    2: {
        "title": "Literature Review",
        "label": "ch:literature-review",
        "file": "chapter2.tex",
    },
    3: {
        "title": "Research Methodology",
        "label": "ch:methodology",
        "file": "chapter3.tex",
    },
}


def convert_word_to_latex(word_file: Path) -> Optional[str]:
    """Convert Word document to LaTeX using Pandoc."""
    print(f"[INFO] Converting {word_file.name} to LaTeX...")
    
    try:
        result = subprocess.run(
            [
                "pandoc",
                str(word_file),
                "-f", "docx",
                "-t", "latex",
                "--wrap=none",
            ],
            capture_output=True,
            check=True,
            encoding='utf-8',
            errors='replace',
        )
        return result.stdout
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] Pandoc conversion failed: {e.stderr}")
        return None
    except FileNotFoundError:
        print("[ERROR] Pandoc not found. Please install Pandoc: https://pandoc.org/installing.html")
        return None


def fix_special_characters(content: str) -> str:
    """Fix special characters and formatting."""
    # Fix tildes (approximate symbol)
    content = re.sub(r'\\textasciitilde\s*', r'$\\sim$', content)
    content = re.sub(r'~(?=\d)', r'$\\sim$', content)
    
    # Fix degree symbol
    content = re.sub(r'°', r'\\textdegree{}', content)
    
    # Fix subscripts in chemical formulas (CO2, CH4, etc.)
    content = re.sub(r'\bCO2\b', r'CO$_2$', content)
    content = re.sub(r'\bCH4\b', r'CH$_4$', content)
    content = re.sub(r'\bH2O\b', r'H$_2$O', content)
    content = re.sub(r'\bO2\b', r'O$_2$', content)
    
    # Fix multiplication symbol
    content = re.sub(r'×', r'$\\times$', content)
    
    # Fix quotes - convert straight quotes to LaTeX curly quotes
    content = re.sub(r'"([^"]*)"', r"``\\1''", content)
    
    # Fix em-dash and en-dash
    content = re.sub(r'—', r'---', content)
    content = re.sub(r'–', r'--', content)
    
    # Fix percent symbol (escape if not already)
    content = re.sub(r'(?<!\\)%', r'\\%', content)
    
    # Fix ampersand (escape if not already)
    content = re.sub(r'(?<!\\)&(?!\\)', r'\\&', content)
    
    return content


def replace_acronyms(content: str) -> str:
    """Replace acronym patterns with \\ac{} commands."""
    used_acronyms = set()
    
    for short, full in ACRONYMS.items():
        # Pattern: "Full Form (SHORT)" -> \ac{SHORT}
        pattern = rf'{re.escape(full)}\s*\({re.escape(short)}\)'
        if re.search(pattern, content, re.IGNORECASE):
            content = re.sub(pattern, rf'\\ac{{{short}}}', content, flags=re.IGNORECASE)
            used_acronyms.add(short)
        
        # Pattern: standalone SHORT (word boundary) that hasn't been replaced
        # Only replace if the full form was seen earlier (i.e., acronym was introduced)
        if short in used_acronyms:
            # Don't replace if already in \ac{} command
            pattern = rf'(?<!\\ac\{{)(?<![a-zA-Z]){re.escape(short)}(?![a-zA-Z\}}])'
            content = re.sub(pattern, rf'\\ac{{{short}}}', content)
    
    return content


def fix_headings(content: str) -> str:
    """Convert Pandoc headings and bold headings to proper LaTeX sectioning commands."""
    
    # First, insert Chapter 1 before "Background and Significance" (first real content)
    # This handles the case where the Word doc doesn't have an explicit Chapter 1 header
    if 'Background and Significance' in content and '\\chapter{Introduction}' not in content:
        content = re.sub(
            r'Background and Significance',
            r'\\chapter{Introduction}\n\\label{ch:introduction}\n\n\\section{Background and Significance}\n\\label{sec:background-and-significance}',
            content,
            count=1
        )
    
    # Handle "Chapter 2" followed by "Literature review" pattern
    content = re.sub(
        r'Chapter\s+2\s*\n+\s*Literature\s+review',
        r'\\chapter{Literature Review}\n\\label{ch:literature-review}',
        content,
        flags=re.IGNORECASE
    )
    
    # Handle "Chapter 3" with "3. Research Methodology" pattern
    content = re.sub(
        r'\\textbf\{Chapter\s+3\}\s*\n+\s*\\textbf\{3\.\s*Research\s+Methodology\}',
        r'\\chapter{Research Methodology}\n\\label{ch:methodology}',
        content,
        flags=re.IGNORECASE
    )
    
    # Also handle plain text Chapter 3 pattern
    content = re.sub(
        r'Chapter\s+3\s*\n+.*?Research\s+Methodology',
        r'\\chapter{Research Methodology}\n\\label{ch:methodology}',
        content,
        flags=re.IGNORECASE
    )
    
    # Section-level headings (bold standalone lines that are section titles)
    SECTION_PATTERNS = [
        # Chapter 1 sections  
        "Developing Circular Economies",
        "Meeting Global Demand and Feedstock Diversification",
        "Emerging Technological Interest",
        "Problem Statement and Research Motivation",
        "Research Objectives and Scope",
        # Chapter 2 sections
        "Molecular Structure and Colloidal Stabilization",
        "Colloids",
        "The Theory of Natural Rubber Latex",
        "Suspensions Rheology and Theoretical Models",
        "Additive Manufacturing",
        # Chapter 3 sections
        "Sourcing, Traceability, and Rationale",
        "Spectroscopic and Analytical Reagents",
        "Rheological Characterization",
        "Nuclear Magnetic Resonance",
        "Components and Protocols",
        "Material Characterizations",
        "Mechanical Characterizations",
        "DLP 3D Printing",
    ]
    
    # Subsection patterns
    SUBSECTION_PATTERNS = [
        "Preservation Chemistry",
        "Preservation-aware framework",
        "Primary Objectives",
        "Research scope",
        "Natural Rubber Latex",
        "Alternative Preservation Systems",
        "Surfactant Stabilization Mechanisms",
        "Structure-Process-Property",
        "Predictive Viscosity Models",
        "Yield Stress and Shear-Thinning",
        "Photoresins for Vat Photopolymerization",
        "Challenges of Elastomers",
        "Polyurethane and Silicone Elastomers",
        "Emulsion-Based 3D Printing",
        "Overview of Preservation Methods",
        "Ammoniated Latex System",
        "Eco-preserved Latex",
        "Reference and Synthetic Polymer",
        "Pre-receipt Specifications",
        "Incoming Verification",
        "Deuterated Solvents",
        "Photopolymerization and Surfactant",
        "Sample Preparation",
        "Rotational Rheometry",
        "Rheology Interpretation",
        "Viscosity and Volume Fraction",
        "Overview and General Conditions",
        "High-Resolution Solution-State",
        "Diffusion-Ordered Spectroscopy",
        "Time-Domain NMR",
        "Preparation of the UV-Curable",
        "Preparation of Photoresin Emulsion",
        "Preparation of Dual Emulsion",
        "Particle Size and Zeta Potential",
        "Photorheology Characterizations",
        "Uniaxial Tension Test",
        "Cyclic and Hysteresis",
        "Fracture Energy",
        "Puncture Tests",
        "Measurement of Curing Depth",
        "Gel Permeation Chromatography",
    ]
    
    # Convert \textbf{Section Title} on its own line to \section{}
    for section_title in SECTION_PATTERNS:
        # Match \textbf{...title...} where title contains the section name
        pattern = rf'\\textbf\{{{re.escape(section_title)}[^}}]*\}}'
        label = section_title.lower().replace(' ', '-').replace(':', '').replace(',', '')
        label = re.sub(r'[^a-z0-9-]', '', label)
        replacement = f'\\\\section{{{section_title}}}\n\\\\label{{sec:{label}}}'
        content = re.sub(pattern, replacement, content, flags=re.IGNORECASE)
    
    # Convert \textbf{Subsection Title} to \subsection{}
    for subsection_title in SUBSECTION_PATTERNS:
        pattern = rf'\\textbf\{{{re.escape(subsection_title)}[^}}]*\}}'
        label = subsection_title.lower().replace(' ', '-').replace(':', '').replace(',', '')
        label = re.sub(r'[^a-z0-9-]', '', label)
        replacement = f'\\\\subsection{{{subsection_title}}}\n\\\\label{{subsec:{label}}}'
        content = re.sub(pattern, replacement, content, flags=re.IGNORECASE)
    
    # Handle any remaining \textbf{} that looks like a heading (standalone on a line)
    lines = content.split('\n')
    result = []
    
    for i, line in enumerate(lines):
        stripped = line.strip()
        
        # Check for hypertarget wrapped headings (Pandoc style)
        hyper_match = re.match(
            r'\\hypertarget\{[^}]*\}\{%?\s*\\(chapter|section|subsection|subsubsection)\{([^}]*)\}[^}]*\}',
            stripped
        )
        if hyper_match:
            cmd = hyper_match.group(1)
            title = hyper_match.group(2)
            label = title.lower().replace(' ', '-').replace(':', '')
            label = re.sub(r'[^a-z0-9-]', '', label)
            result.append(f'\\{cmd}{{{title}}}')
            if cmd == 'chapter':
                result.append(f'\\label{{ch:{label}}}')
            else:
                result.append(f'\\label{{sec:{label}}}')
            continue
        
        # Check if line is just \textbf{Some Title} (potential missed heading)
        bold_only = re.match(r'^\\textbf\{([^}]+)\}$', stripped)
        if bold_only:
            title = bold_only.group(1)
            # If title looks like a heading (short, capitalized words)
            if len(title) < 100 and not title.endswith('.'):
                label = title.lower().replace(' ', '-').replace(':', '')
                label = re.sub(r'[^a-z0-9-]', '', label)
                # Determine level based on context or default to subsection
                result.append(f'\\subsection*{{{title}}}')
                result.append(f'\\label{{subsec:{label}}}')
                continue
        
        result.append(line)
    
    return '\n'.join(result)


def split_into_chapters(content: str) -> dict[int, str]:
    """Split content into separate chapters."""
    chapters = {}
    
    # Find chapter boundaries
    chapter_pattern = r'\\chapter\{([^}]*)\}'
    matches = list(re.finditer(chapter_pattern, content))
    
    if not matches:
        print("[WARNING] No \\chapter{} commands found. Content may need manual review.")
        return {1: content}
    
    # Remove frontmatter (everything before first chapter)
    first_chapter_start = matches[0].start()
    frontmatter = content[:first_chapter_start].strip()
    if frontmatter:
        print(f"[INFO] Removed {len(frontmatter)} characters of frontmatter before Chapter 1")
    
    for i, match in enumerate(matches):
        chapter_num = i + 1
        start = match.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(content)
        
        chapter_content = content[start:end].strip()
        chapters[chapter_num] = chapter_content
    
    return chapters


def create_chapter_file(chapter_num: int, content: str) -> str:
    """Create a complete chapter file with header."""
    config = CHAPTERS.get(chapter_num, {
        "title": f"Chapter {chapter_num}",
        "label": f"ch:chapter{chapter_num}",
        "file": f"chapter{chapter_num}.tex",
    })
    
    header = f"""% {config['file']} -- Chapter {chapter_num}: {config['title']}
%
% Natural Rubber Latex Thesis
% Auto-generated from Word document - {time.strftime('%Y-%m-%d %H:%M:%S')}

"""
    return header + content


def process_and_save(content: str) -> None:
    """Process LaTeX content and save to chapter files."""
    # Apply all transformations
    content = fix_special_characters(content)
    content = replace_acronyms(content)
    content = fix_headings(content)
    
    # Split into chapters
    chapters = split_into_chapters(content)
    
    # Save each chapter
    for chapter_num, chapter_content in chapters.items():
        if chapter_num not in CHAPTERS:
            print(f"[WARNING] Chapter {chapter_num} not in configuration, skipping.")
            continue
        
        config = CHAPTERS[chapter_num]
        output_file = CHAPTERS_DIR / config["file"]
        
        full_content = create_chapter_file(chapter_num, chapter_content)
        
        # Backup existing file
        if output_file.exists():
            backup_file = output_file.with_suffix('.tex.bak')
            backup_file.write_text(output_file.read_text(encoding='utf-8'), encoding='utf-8')
        
        output_file.write_text(full_content, encoding='utf-8')
        print(f"[OK] Saved {output_file.name}")


def convert_once() -> bool:
    """Perform a single conversion."""
    if not WORD_FILE.exists():
        print(f"[ERROR] Word file not found: {WORD_FILE}")
        return False
    
    content = convert_word_to_latex(WORD_FILE)
    if content is None:
        return False
    
    process_and_save(content)
    print("[OK] Conversion complete!")
    return True


def watch_and_convert() -> None:
    """Watch for changes and convert automatically."""
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
    
    class WordFileHandler(FileSystemEventHandler):
        def __init__(self):
            self.last_modified = 0
            self.debounce_seconds = 2  # Wait 2 seconds after last change
        
        def on_modified(self, event):
            if event.is_directory:
                return
            
            # Check if it's our Word file
            if Path(event.src_path).resolve() == WORD_FILE.resolve():
                current_time = time.time()
                if current_time - self.last_modified > self.debounce_seconds:
                    self.last_modified = current_time
                    print(f"\n[DETECTED] Change in {WORD_FILE.name}")
                    time.sleep(1)  # Wait for file to be fully written
                    convert_once()
    
    # Initial conversion
    print("[INFO] Performing initial conversion...")
    convert_once()
    
    # Set up watcher
    event_handler = WordFileHandler()
    observer = Observer()
    observer.schedule(event_handler, str(WORD_FILE.parent), recursive=False)
    observer.start()
    
    print(f"\n[WATCHING] Monitoring {WORD_FILE}")
    print("[INFO] Press Ctrl+C to stop...")
    
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        print("\n[INFO] Watcher stopped.")
    
    observer.join()


def main():
    """Main entry point."""
    print("=" * 60)
    print("Word-to-LaTeX Thesis Converter")
    print("=" * 60)
    
    if "--once" in sys.argv:
        success = convert_once()
        sys.exit(0 if success else 1)
    else:
        watch_and_convert()


if __name__ == "__main__":
    main()
