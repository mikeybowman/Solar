# Solar - The Open Source Game Design Document Builder

Solar is a desktop application for creating and maintaining Game Design Documents without relying on online services, proprietary tools, or rigid templates. It's designed for developers who want structured documentation that scales with a project, while still remaining easy to edit and version control.

Solar began as an internal tool built at Dead Orbit Studios to manage large, evolving GDDs. It has since been open-sourced for anyone who wants a practical, no-nonsense documentation workflow. While Solar will get you the majority of the way there, if you want a really refined and presentation ready GDD, you'll most certainly want to go in and reformat the exported document to suit your unique needs.

## Overview

Solar focuses on three core ideas:

- Keep documentation local and human-readable  
- Make large design documents manageable  
- Support real production workflows, including version control  

The application uses a simple desktop interface with collapsible entries to prevent large projects from becoming unreadable walls of text. Version 2.0.0 introduces three specialized professional documentation tabs (Combat Mechanics, Player Progression, and Map Design) that expand Solar's capabilities to over 170 comprehensive documentation fields, suitable for documenting complex game systems at professional development standards.

## Features

### Structured GDD Editing

- Static Introduction section for high-level project information and development timeline
- **Design Pillars section** for documenting core design principles and guidelines
- Dynamic sections for mechanics, levels, characters, items, audio, art, and technical design
- Collapsible entries to keep large documents navigable
- Entry headers update automatically based on name fields
- Comprehensive field coverage ensures no data is lost during export

<table align="center">
  <tr>
    <td align="center">
      <img src="assets/intro.png" width="300"><br>
      Intro Example
    </td>
    <td align="center">
      <img src="assets/closed.png" width="300"><br>
      Collapsed Entries
    </td>
    <td align="center">
      <img src="assets/open.png" width="300"><br>
      Expanded Entries
    </td>
  </tr>
</table>

### Design Pillars Management

Solar includes a dedicated Design Pillars section to help teams maintain focus on core design principles:

- **Pillar Name**: Clear identification of each design principle
- **Core Principle**: The fundamental concept driving decisions
- **Why It Matters**: Importance and impact on the game experience
- **Implementation Guidelines**: Practical guidance for the development team
- **Examples in Game**: Concrete examples of the pillar in action

This ensures all team members understand and can refer back to the foundational design decisions throughout development.

### Documentation Tabs

Version 2.0.0 introduces three specialized tabs designed for professional game development workflows:

**Combat Mechanics Documentation**
Solar now includes comprehensive documentation for weapons, abilities, and combat systems with 27 professional fields:

- **Core Statistics**: Damage values, rate of fire, reload times, ammunition capacity, and range metrics
- **Balance Analysis**: Strengths, weaknesses, counter-play options, competitive usage statistics, and balance history
- **Special Mechanics**: Abilities, status effects, cooldowns, resource costs, and area of effect documentation
- **Technical Implementation**: Known issues tracking, implementation notes, and performance considerations
- **Competitive Data**: Pick rates, win rates, and professional usage analysis

**Player Progression Systems**
Complete documentation for retention mechanics and monetization systems with 27 specialized fields:

- **Progression Mechanics**: XP scaling formulas, level caps, time investment estimates, and unlock requirements
- **Reward Systems**: Milestone rewards, daily bonuses, seasonal events, and special unlock mechanisms
- **Currency Economics**: Earning rates, spending options, monetization impact analysis, and economic balance
- **Retention Strategies**: Engagement hooks, social features, FOMO elements, and player feedback integration
- **Analytics Integration**: Player behavior tracking, retention metrics, and progression analytics

**Map Design Analysis**
Professional level design documentation with 27 fields for strategic and competitive analysis:

- **Layout Specifications**: Dimensions, player counts, spawn systems, and objective placement
- **Strategic Elements**: Sightlines, cover positions, flanking routes, and high ground advantages
- **Communication Systems**: Official callouts, community terminology, and strategic zone definitions
- **Environmental Design**: Interactive hazards, destructible objects, lighting considerations, and atmospheric effects
- **Competitive Analysis**: Traffic flow patterns, balance considerations, known exploits, and tournament viability

These additions bring Solar's total documentation capacity to over 170 professional fields, making it suitable for documenting complex game systems found in AAA development environments.


### Export Options

- Microsoft Word (.docx) export with clean section formatting and complete field coverage
- Plain text export for portability and backups with full data preservation
- JSON project files for native saving and version control
- Structured file export that generates folders and individual `.txt` files per entry
- Consistent author and studio attribution across all export formats

All exports now include every field from every section, ensuring complete data preservation regardless of export format chosen.

The structured export is designed specifically for Git-based workflows, allowing granular diffs and modular collaboration.

### Quality-of-Life Features

- Persistent settings (project folder, window size, author/studio info)
- Cross-platform support (Windows, macOS, Linux)
- Runs as a Python script or standalone Windows executable
- No accounts, no cloud services, fully offline
- Clean, professional exports without repetitive footer credits
- Configurable studio and author information with consistent export formatting
- **Enhanced Interface**: Natural mouse wheel scrolling throughout the application, eliminating the need to find small scrollbars
- **Optimized Window Sizing**: Default window size increased to 1400x800 with minimum size constraints to properly display all professional tabs
- **Extended Session Stability**: Production-grade memory management enabling reliable operation during multi-hour documentation sessions
- **Comprehensive Error Handling**: Enhanced error recovery and detailed logging ensuring stable operation during complex workflows

<p align="center">
  <img src="assets/configure_studio_author.png" width="300">
</p>

## Installation

### Option 1: Windows Executable

1. Download `Solar.exe` from the Releases page
2. Run directly (no installer required)
3. Set a project folder and begin documenting

### Option 2: Run From Source (All Platforms)

```bash
git clone https://github.com/mikeybowman/Solar.git
cd Solar
pip install -r requirements.txt
python gdd_builder.py
```

### Option 3: Build Your Own Executable

```bash
pip install pyinstaller python-docx
build_exe.bat
```
You may also build manually using PyInstaller if you need custom options.

### Basic Workflow

1. Launch Solar and configure studio and author information
2. Set a project folder (remembered between sessions)
3. Fill out the Introduction tab with project details and development timeline
4. Define your core Design Pillars to guide development decisions
5. Add entries to each section as needed (mechanics, levels, characters, etc.)
6. Save progress using the JSON project format
7. Export to Word, text, or structured files when required

### File Structure Export

The structured export mirrors the internal layout of the application:
```bash
YourGame_GDD/
├── README.txt
├── 01_Introduction/
├── 02_Design_Pillars/
├── 03_Mechanics/
├── 04_Levels/
├── 05_Characters/
├── 06_Items/
├── 07_Audio/
├── 08_Art_Style/
└── 09_Technical/
```
Each entry is written to its own file, making it easy to track changes, review diffs, and share specific sections with collaborators.

<p align="center">
  <img src="assets/file_structure.png" width="300">
</p>

### Design Philosophy

Solar intentionally avoids modern web-app design patterns. The interface is simple, predictable, and focused on readability. The collapsible entry system exists to solve a real problem: GDDs grow faster than most tools can handle.

The addition of Design Pillars reflects the reality of professional game development: successful projects maintain clear vision throughout development by documenting and referencing core design principles.

The goal is to stay out of the way and let the documentation speak for itself.

### Development Notes

- GUI: Tkinter (for maximum compatibility)
- Data storage: JSON (human-readable, version-control friendly)
- Word export: python-docx with complete field coverage
- Settings: JSON-based with sane defaults
- Architecture: Modular tab system for easy feature additions

The codebase is intentionally straightforward. Adding new sections or fields follows established patterns and does not require major refactoring.

### Recent Updates (v2.0.0)

**Major New Features:**
- **Three New Professional Documentation Tabs**: Combat Mechanics, Player Progression, and Map Design with 81 additional specialized fields
- **Enhanced User Interface**: Improved window sizing (1400x800 default) with natural mouse wheel scrolling throughout the application
- **Production-Grade Stability**: Comprehensive memory management and error handling for extended documentation sessions
- **Professional Sample Project**: Complete "Void Arena" GDD demonstrating all features with industry-quality content

**Documentation Enhancements:**
- **Expanded Structure**: Complete 12-section documentation framework suitable for AAA development standards
- **Enhanced Export System**: Improved formatting and professional presentation across all export formats
- **Field Coverage**: Total documentation capacity increased from 120+ to 170+ professional fields
- **Cross-Platform Improvements**: Enhanced stability and performance across Windows, macOS, and Linux

**Stability and Performance:**
- **Memory Leak Prevention**: Automatic resource cleanup during extended sessions
- **Enhanced Error Handling**: Graceful recovery from edge cases with comprehensive logging
- **Performance Optimization**: Improved responsiveness for large documentation projects
- **Professional Reliability**: Production-ready stability for commercial development workflows

**Previous Updates (v1.1.0):**
- Design Pillars tab for documenting core design principles
- Complete field export coverage across all categories
- Enhanced author attribution in all export formats
- Fixed missing fields in Word document exports and improved export consistency

### Contributing

Contributions are welcome.

- Bug reports and feature requests can be filed as GitHub issues
- Pull requests should follow existing code structure and style
- Documentation improvements are appreciated

If you fork Solar for your own workflow, feel free to adapt it as needed.

### Use Cases

Solar is suitable for:

- Independent studios maintaining production documentation
- Game design students learning industry-standard structure and design pillar methodology
- Hobby developers organizing long-term projects
- Educators teaching documentation and design planning
- Teams that need to maintain design consistency across development cycles

### Security & Integrity
### SHA-256 Checksums

For released binaries, SHA-256 checksums are provided to verify file integrity.

```bash
Version: 2.0.0
File: Solar.exe
Size: 17929503 bytes
SHA256: 4f55a95205633324861c28cf2254b81155d105a627770a95949d8409d7f3aac5
```
```bash
Version: 1.1.0
File: Solar.exe
Size: 17919078 bytes
SHA256: d2559b67bc01448f7196f826911c2265efc8444713a2baacf9818260fec820d4
```
Always verify checksums before running downloaded executables.

**To generate hash:**
```bash
# Windows
certutil -hashfile Solar.exe SHA256

# Linux/Mac
sha256sum Solar.exe
```

# License
Solar is open source software. See the LICENSE file for details.

# Credits
Original creator: Mikey LaBrecque
Studio: Dead Orbit Studios

Solar exists to make game design documentation easier to manage and easier to maintain. If it helps you organize ideas, maintain design consistency, or ship a project, it's doing its job.
