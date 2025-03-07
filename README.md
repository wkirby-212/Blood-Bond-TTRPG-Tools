# Blood Bond Tools

## Overview

Blood Bond Tools is a suite of utilities designed to aid in the creation and management of dialog spells for the Blood Bond system. The tools provide a user-friendly interface for generating spells with various properties, durations, and elements, making spell creation more efficient and consistent.

## Main Components

### DialogSpellMaker

`DialogSpellMaker.ps1` is a GUI-based PowerShell script that allows you to:

- Create dialog spells with custom properties
- Manage spell durations and effects
- Generate spell data compatible with the Blood Bond system
- Export spell configurations for use in-game

### ElementMapper

`ElementMapper.ps1` is responsible for:

- Mapping elements between different systems
- Ensuring compatibility between spell data and templates
- Converting element types to maintain consistency across the Blood Bond framework

## Installation

1. Download or clone the Blood Bond Tools repository to your computer
2. Navigate to the main directory
3. Run `SHortcut Creator.bat` to create shortcuts and set up the tools
   - This will create desktop shortcuts for easy access to the tools
   - The setup ensures all dependencies are properly connected

## Configuration Files

The tools rely on several JSON configuration files:

- **spelling_descriptions.json**: Contains detailed descriptions for various spells
- **spelling_synonyms.json**: Provides alternative terms and synonyms for spell components
- **spell_timing_patterns.json**: Defines timing patterns and duration templates for spells
- Additional JSON files store spell data, element mappings, and other configuration information

## Usage

### Creating a New Spell

1. Launch DialogSpellMaker using the desktop shortcut or by running `DialogSpellMaker.ps1`
2. Select the desired spell properties from the GUI
3. Configure duration, elements, and other attributes
4. Generate the spell data
5. Export or save the configuration for use in-game

### Element Mapping

1. Run ElementMapper when you need to convert or map elements between systems
2. Select the source and target element systems
3. Follow the prompts to complete the mapping process

## Requirements

- Windows operating system
- PowerShell 5.0 or higher
- Administrative privileges (for initial setup)

## Support

For issues, questions, or contributions, please contact the project maintainer.

---

*Blood Bond Tools - Enhancing your magical experience*

