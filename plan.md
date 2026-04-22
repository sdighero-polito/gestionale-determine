# Auto

## Configuration
- **Artifacts Path**: .zenflow/tasks/1b43bf91-aa0e-4207-a97f-142fe433fe08

### [x] Step 1: Investigation and File Organization
- Renamed attachment files to correct names (app.py, database.py, etc.)
- Moved files to project root

### [x] Step 2: Code Refactoring and Improvements
- Centralized configuration: moved hardcoded strings (Ente, Dirigente, etc.) to database settings
- Added management UI for general settings (Settings window)
- Improved database layer with proper update functions for Verbali and Verifiche
- Enhanced document generation to use dynamic settings
- Added Database Backup functionality to the UI

### [x] Step 3: Fix Financial Coverage and QE in Documents
- [x] Fix financial coverage (coperture finanziarie) saving logic
- [x] Implement Quadro Economico (QE) table generation in Word documents (Determina)
- [x] Add "AI Guidance" field to allow users to direct AI content generation

### [x] Step 5: Registry Simplification and Coverage Fixes
- [x] Removed fixed roles from Personnel Registry (defined per-assignment instead)
- [x] Added total display to Financial Coverage tab
- [x] Made coverage fields optional (Capitolo no longer mandatory in DB)
