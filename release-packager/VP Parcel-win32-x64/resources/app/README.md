# VP Parcel - Shipping Label Generator

VP Parcel is a streamlined desktop application designed to automate the creation of shipping labels from Excel data. Built with Electron and React, it provides a fast and professional way to generate bulk PDF labels.

## 🚀 Features

- **Excel Data Import**: Easily upload `.xlsx` or `.xls` files to process bulk shipping data.
- **Automated PDF Generation**: Automatically formats and generates high-quality shipping labels.
- **Custom Layout**: Labels include formatted sender and receiver details, amount in words, and tracking dimensions.
- **Bulk Processing**: Handles multiple entries from a single sheet efficiently.

## 🛠️ Tech Stack

- **Framework**: [Electron](https://www.electronjs.org/) (Desktop Application)
- **Frontend**: [React](https://react.dev/) + [Vite](https://vitejs.dev/)
- **Libraries**:
  - `exceljs`: For parsing Excel spreadsheets.
  - `jspdf` & `jspdf-autotable`: For generating and formatting PDF documents.

## 📦 Installation & Setup

1. **Clone the repository**:
   ```bash
   git clone https://github.com/Vikranth-1/VP_Parcel.git
   cd "PROJECT FILES"
   ```

2. **Install dependencies**:
   ```bash
   npm install
   ```

3. **Run in development mode**:
   ```bash
   npm run electron:dev
   ```

4. **Build for production**:
   ```bash
   npm run electron:build
   ```

## 📂 Project Structure

- `src/`: React frontend components and logic.
- `main.cjs`: Electron main process configuration.
- `preload.cjs`: Electron preload script for secure inter-process communication.
- `public/`: Static assets.

## 📄 License

This project is for internal use for shipping management.
