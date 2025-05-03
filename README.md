# VAD Prescription System

A specialized medical prescription management system designed for VAD (Ventricular Assist Device) outpatient clinics. This application streamlines the process of creating standardized test prescriptions for patient follow-up visits.

## Features

- **Patient Management**: Import patient data from Excel spreadsheets
- **Prescription Management**: Create, save, and organize test prescriptions
- **Standardized Tests**: Common blood tests, additional tests, and imaging tests
- **Custom Tests**: Add and manage custom examinations
- **Test Groups**: Create and save groups of tests for quick application
- **Document Generation**: Export prescriptions to formatted Word documents
- **Data Persistence**: Automatically save settings and configurations between sessions

## Installation

### Requirements

- Python 3.6+
- Required packages:
  - python-docx
  - pandas
  - openpyxl
  - tkcalendar
  - tkinter (usually included with Python)

### Installation Steps

You can download the executables from the Releases for Windows and Ubuntu

On Ubuntu or other Linux distributions:

Make it executable:
   ```bash
   chmod 755 VAD_prescription_scheduler_{version number}-ubuntu latest
   ```

and then you can lauch it:
   ```bash
   ./VAD_prescription_scheduler_{version number}-ubuntu latest
   ```

where {version number} is the version of the release you downloaded.

## Usage

### Input Data Format

The system expects an Excel file with at least the following columns:
1. A.M. (Hospital ID)
2. Επώνυμο (Last Name)
3. Όνομα (First Name)
4. ΑΜΚΑ (Social Security Number)
5. Απόφαση ΚΕΠΑ (KEPA Decision)

### Basic Workflow

1. **Load Patient Data**:
   - Click "Αναζήτηση..." to select an Excel file with patient data
   - The system will remember the last used file and can load it automatically

2. **Set the Date**:
   - Use the date picker to select the desired date for the prescription

3. **Select a Patient**:
   - Choose a patient from the list on the left side

4. **Select Tests**:
   - Check the desired tests in the categories on the right side
   - Apply test groups if needed
   - Add notes in the text field

5. **Apply Selections**:
   - Click "Εφαρμογή Επιλογών" to save the selections for the current patient

6. **Repeat for All Patients**:
   - Select the next patient and repeat steps 4-5

7. **Generate Document**:
   - Click "Δημιουργία Εγγράφου" to create a Word document with all prescriptions
   - Select a location to save the document

### Managing Test Groups

1. **Open Group Manager**:
   - Click "Διαχείριση Ομάδων Εξετάσεων"

2. **Create New Group**:
   - Click "Δημιουργία Νέας Ομάδας"
   - Enter a name for the group
   - Select the tests to include in the group
   - Click "Αποθήκευση"

3. **Apply Group to Patient**:
   - Select a patient
   - Choose a group from the dropdown
   - Click "Εφαρμογή Ομάδας"

## Data Storage

The application stores several types of data in JSON files:
- `custom_exams.json`: Custom examination definitions
- `exam_groups.json`: Saved test groups
- `settings.json`: Application settings, including the last used file path

## Output Format

The generated Word document includes:
- Header with the clinic name and date
- Table with patient information
- Selected tests grouped by category
- Patient notes

## License

This software is distributed under the MIT License.

## Acknowledgments

Developed for the VAD outpatient clinic in my centre to improve workflow efficiency and standardize test prescriptions.
