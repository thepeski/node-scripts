/* Excel.js */
// manipulating excel workbooks

// default imports
import ExcelJS from "exceljs";

export class Excel {
    constructor() {
        this.isOpen = false;
        this.inputDir = null;

        this.workbooks = {};
        this.activeWorkbook = null;
        this.activeSheet = null;
    }

    /* file management */

    async open(files, directory = "./", extension = "xlsx") {

        // remember directory for quick save
        this.inputDir = directory;

        // ensure files is loopable
        if (!Array.isArray(files)) files = [files];

        // loop over files
        for (const file of files) {
            const path = `${directory}/${file}.${extension}`;

            // initialise workbook
            const workbook = new ExcelJS.Workbook();

            try {
                await workbook.xlsx.readFile(path);

                // add to workbooks object & flip indicator
                this.workbooks[file] = workbook;
                this.isOpen = true;

                console.log("opened:", `${file}.${extension}`);
            } catch (error) {
                console.error("opening failed:", error.message || String(error));

                return false;
            }

        }

        // select workbook if only one opened & loaded
        if (files.length === 1 && Object.keys(this.workbooks).length === 1) {
            this.activeWorkbook = files[0];

            console.log("using workbook:", files[0]);
        }

        return true;
    }

    async save(directory, extension = "xlsx") {

        // no active workbook
        if (!this.activeWorkbook) {
            console.error("no workbook selected");

            return false;
        }

        // use provided directory or input directory, default to root
        const outputDir = directory ?? this.inputDir ?? "./";

        const path = `${outputDir}/${this.activeWorkbook}.${extension}`;

        try {
            await this.workbooks[this.activeWorkbook].xlsx.writeFile(path);
        } catch (error) {
            console.error("saving failed:", error.message || String(error));
        }

        console.log("saved:", `${this.activeWorkbook}.${extension}`);

        return true;
    }

    async saveAs(files, directory, extension = "xlsx") {

        // no active workbook
        if (!this.activeWorkbook) {
            console.error("no workbook selected");

            return false;
        }

        // use provided directory or input directory, default to root
        const outputDir = directory ?? this.inputDir ?? "./";

        // ensure files is loopable
        if (!Array.isArray(files)) files = [files];

        // loop over files
        for (const file of files) {
            const path = `${outputDir}/${file}.${extension}`;

            try {
                await this.workbooks[this.activeWorkbook].xlsx.writeFile(path);

                console.log("saved as:", `${file}.${extension}`);
            } catch (error) {
                console.error(
                    `saving as ${file}.${extension} failed:`,
                    error.message || String(error)
                );
            }
        }

        return true;
    }

    async saveAll(directory, extension = "xlsx") {

        // no workbooks open
        if (!this.isOpen) {
            console.error("no workbooks open");

            return false;
        }

        // use provided directory or input directory, default to root
        const outputDir = directory ?? this.inputDir ?? "./";

        // loop over object keys
        for (const workbook of Object.keys(this.workbooks)) {
            const path = `${outputDir}/${workbook}.${extension}`;

            try {
                await this.workbooks[workbook].xlsx.writeFile(path);

                console.log("saved:", `${workbook}.${extension}`);
            } catch (error) {
                console.error(
                    `saving ${workbook}.${extension} failed:`,
                    error.message || String(error)
                );

                return false;
            }
        }

        return true;
    }

    close(workbook) {

        const workbookToClose = workbook ?? this.activeWorkbook;

        // no active workbook
        if (!workbookToClose) {
            console.error("no workbook selected to close");

            return false;
        }

        // remove property from workbooks object
        delete this.workbooks[workbookToClose];

        console.log("closed:", workbookToClose);

        // reset
        this.activeWorkbook = null;
        this.activeSheet = null;

        // flip indicator & clear input directory if no workbooks left open
        if (Object.keys(this.workbooks).length === 0) {
            this.isOpen = false;
            this.inputDir = null;

            console.log("no workbooks left open");
        }

        return true;
    }

    closeAll() {

        // reset
        this.activeWorkbook = null;
        this.activeSheet = null;

        // loop over object keys & delete
        for (const workbook of Object.keys(this.workbooks)) {
            delete this.workbooks[workbook];

            console.log("closed:", workbook);
        }

        if (Object.keys(this.workbooks).length === 0) {
            this.isOpen = false;
            this.inputDir = null;

            console.log("no workbooks left open");
        }

        return true;
    }

    /* workbooks */

    listWorkboks() { }

    addWorkbook() { }

    useWorkbook() { }

    deleteWorkbook() { }

    getActiveWorkbook() { }

    getActiveWorkbookName() { }

    /* sheets */

    listSheets() { }

    addSheet() { }

    useSheet() { }

    deleteSheet() { }

    getActiveSheet() { }

    getActiveSheetName() { }

    /* fetching data */

    fetchCell() { }

    fetchCells() { }

    fetchValue() { }

    fetchValues() { }

    fetchFormula() { }

    fetchFormulas() { }

    /* setting data */

    setCell() { }

    setCells() { }

    setValue() { }

    setValues() { }

    setFormula() { }

    setFormulas() { }

    /* clearing data */

    clear() { }

    /* copying data */

    copyCell() { }

    copyCells() { }

    copyValue() { }

    copyValues() { }

    copyFormula() { }

    copyFormulas() { }

    /* utilities */

    fillValue() { }

    fillFormula() { }

    /* styling */

    fetchStyle() { }

    fetchStyles() { }

    setStyle() { }

    setStyles() { }
}