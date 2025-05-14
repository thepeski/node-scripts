/* Excel.js */
// editing excel workbooks

// default imports
import ExcelJS from "exceljs";
import fs from "fs/promises";

export class Excel {
    constructor() {
        this.isOpen = false;
        this.inputDir = null;

        this.workbooks = {};
        this.activeWorkbook = null;
        this.activeSheet = null;
    }

    /* file management */

    // open files from directory & load into workbooks object
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
            this.useWorkbook(files[0]);
        }

        return true;
    }

    // save active workbook in place or in a given directory
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

    // save active workbook using specified names in place or in a given directory
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

    // save all loaded workbooks in place or in a given directory
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

    // delete active or selected workbook file in place or in a given directory, optionally close
    async delete(name, close = false, directory = null, extension = "xlsx") {

        // no workbook to delete
        if (!name && !this.activeWorkbook) {
            console.error("no workbook to delete");

            return false;
        }

        // use provided name or active workbook
        const file = name ?? this.activeWorkbook;

        // remove property from workbooks object if close true & property exists
        if (close && Object.keys(this.workbooks).includes(file)) {
            delete this.workbooks[file];

            console.log("closed:", file);
        }

        // use provided directory or input directory, default to root
        const deleteDir = directory ?? this.inputDir ?? "./";

        const path = `${deleteDir}/${file}.${extension}`;

        try {
            await fs.unlink(path);

            console.log("deleted workbook:", file);
        } catch (error) {
            console.error("failed to delete workbook:", path);

            return false;
        }

        return true;
    }

    // close active or selected workbook
    close(workbook) {

        // no workbooks open
        if (!this.isOpen) {
            console.error("no workbooks open");

            return false;
        }

        const workbookToClose = workbook ?? this.activeWorkbook;

        // no active workbook
        if (!workbookToClose) {
            console.error("no workbook selected to close");

            return false;
        }

        // workbook not open
        if (!Object.keys(this.workbooks).includes(workbookToClose)) {
            console.error(`workbook ${workbookToClose} not open`);

            return false;
        }

        // remove property from workbooks object
        delete this.workbooks[workbookToClose];

        // reset properties as active workbook was closed
        if (workbookToClose === this.activeWorkbook) {
            this.activeWorkbook = null;
            this.activeSheet = null;
        }

        // flip indicator & clear input directory if no workbooks left open
        if (Object.keys(this.workbooks).length === 0) {
            this.isOpen = false;
            this.inputDir = null;

            console.log("no workbooks left open");
        }

        return true;
    }

    // close all loaded workbooks
    closeAll() {

        // no workbooks open
        if (!this.isOpen) {
            console.error("no workbooks open");

            return false;
        }

        // reset properties
        this.activeWorkbook = null;
        this.activeSheet = null;

        // loop over object keys & delete
        for (const workbook of Object.keys(this.workbooks)) {
            delete this.workbooks[workbook];

            console.log("closed:", workbook);
        }

        // check if all workbooks were closed
        if (Object.keys(this.workbooks).length === 0) {
            this.isOpen = false;
            this.inputDir = null;

            console.log("no workbooks left open");
        }

        return true;
    }

    /* workbooks */

    // list open workbooks
    listWorkbooks() {

        // no workbooks open
        if (!this.isOpen) {
            console.error("no workbooks open");

            return false;
        }

        console.log("open workbooks:");

        // loop over workbooks
        for (const workbook of Object.keys(this.workbooks)) {
            console.log(workbook);
        }

        return true;
    }

    // add & load new workbook, use if selected
    addWorkbook(name, use = false) {

        // initialise workbook & sheet
        const workbook = new ExcelJS.Workbook();
        const worksheet = "Sheet1";

        // ensure workbook has at least one sheet
        workbook.addWorksheet(worksheet);

        // add to workbooks object
        this.workbooks[name] = workbook;

        // flip indicator if no workbooks open so far
        if (!this.isOpen) this.isOpen = true;

        // set as active workbook & sheet if true
        if (use) {
            this.activeWorkbook = name;
            this.activeSheet = worksheet;
        }

        return true;
    }

    // select given workbook & its first sheet
    useWorkbook(name) {

        // no workbooks open
        if (!this.isOpen) {
            console.error("no workbooks open");

            return false;
        }

        // selected workbook not open
        if (!Object.keys(this.workbooks).includes(name)) {
            console.error(`workbook ${name} not open`);

            return false;
        }

        if (name === this.activeWorkbook) {
            console.log("already using workbook:", name);

            return true;
        }

        // set active workbook name
        this.activeWorkbook = name;

        // check if sheets exist and select the first one
        if (this.workbooks[name].worksheets.length > 0) {
            const sheet = this.workbooks[name].worksheets[0].name;
            this.activeSheet = sheet;

            console.log(`using: workbook ${name} & sheet ${sheet}`);
        } else {
            this.activeSheet = null;

            console.log(`using workbook ${name}`);
        }

        return true;
    }

    // return active or selected workbook object
    getWorkbook(name) {

        // use provided name or active workbook
        const workbook = name ?? this.activeWorkbook;

        // no workbook to return
        if (!workbook) {
            console.error("no workbook to return");

            return false;
        }

        // workbook with given name not open
        if (!Object.keys(this.workbooks).includes(workbook)) {
            console.error(`workbook ${name} not open`);

            return false;
        }

        console.log("returned workbook:", workbook);

        return this.workbooks[workbook];
    }

    // return active workbook name
    getActiveWorkbookName() {

        // no workbook selected
        if (!this.activeWorkbook) {
            console.error("no workbook selected");

            return false;
        }

        console.log("active workbook:", this.activeWorkbook);

        return true;
    }

    /* sheets */

    // list sheets within active or provided workbook
    listSheets(name) {

        // use provided name or active workbook
        const workbook = name ?? this.activeWorkbook;

        // no workbooks to look for sheets in
        if (!workbook) {
            console.error("no workbook to examine");

            return false;
        }

        // workbook with given name not open 
        if (!Object.keys(this.workbooks).includes(workbook)) {
            console.error(`workbook ${workbook} not open`);

            return false;
        }

        console.log("existing sheets:");

        // loop over sheets & list each one
        for (const sheet of this.workbooks[workbook].worksheets) {
            console.log(sheet?.name);
        }

        return true;
    }

    // add sheet to active or provided workbook, use if selected
    addSheet(name, workbook, use = false) {

        // use provided name or active workbook
        const affectedWorkbook = workbook ?? this.activeWorkbook;

        // no workbook to add a sheet to
        if (!affectedWorkbook) {
            console.error("no workbook to add a sheet to");

            return false;
        }

        // given workbook not open
        if (affectedWorkbook && !Object.keys(this.workbooks).includes(affectedWorkbook)) {
            console.error(`workbook ${affectedWorkbook} not open`);

            return false;
        }

        try {
            this.workbooks[affectedWorkbook].addWorksheet(name);
        } catch (error) {
            console.error("failed adding sheet:", error.message || String(error));

            return false;
        }

        console.log(`added sheet ${name} to ${affectedWorkbook}`);

        // set as active sheet if true
        if (use) {
            this.activeSheet = name;

            console.log("using sheet:", name);
        }

        return true;
    }

    // select given sheet from active workbook
    useSheet(name) {

        // no workbook selected
        if (!this.activeWorkbook) {
            console.error("no workbook selected");

            return false;
        }

        const sheets = this.workbooks[this.activeWorkbook].worksheets;

        // sheet does not exist
        if (!sheets.map((s) => s.name).includes(name)) {
            console.error(`sheet ${name} does not exist`);

            return false;
        }

        this.activeSheet = name;

        console.log("using sheet:", name);

        return true;
    }

    // delete sheet from active or provided workbook
    deleteSheet(name, workbook) {

        // use provided name or active workbook
        const affectedWorkbook = workbook ?? this.activeWorkbook;

        // no workbook to delete a sheet from
        if (!affectedWorkbook) {
            console.error("no workbook to delete a sheet from");

            return false;
        }

        // given workbook not open
        if (affectedWorkbook && !Object.keys(this.workbooks).includes(affectedWorkbook)) {
            console.error(`workbook ${affectedWorkbook} not open`);

            return false;
        }

        const sheet = this.workbooks[affectedWorkbook].getWorksheet(name);

        // if sheet exists
        if (sheet) {
            this.workbooks[affectedWorkbook].removeWorksheet(sheet?.id);

            console.log("removed sheet:", sheet?.name);
        } else {
            console.error(`sheet ${name} not found`);

            return false;
        }

        // if deleted active sheet
        if (name === this.activeSheet) {

            // reset property
            this.activeSheet = null;

            console.log("deleted active sheet");
        }

        return true;
    }

    // return active or provided sheet from active or provided workbook
    getSheet(name, workbook) {

        // use provided names or active properties
        const sheet = name ?? this.activeSheet;
        const affectedWorkbook = workbook ?? this.activeWorkbook;

        // no sheet or workbook to use
        if (!sheet) {
            console.error("no sheet to return");

            return false;
        } else if (!affectedWorkbook) {
            console.error("no workbook to return a sheet from");

            return false;
        }

        if (!Object.keys(this.workbooks).includes(affectedWorkbook)) {
            console.error(`workbook ${affectedWorkbook} not open`);

            return false;
        }

        const resultSheet = this.workbooks[affectedWorkbook].getWorksheet(sheet);

        // sheet does not exist
        if (!resultSheet) {
            console.error(`sheet ${sheet} does not exist`);

            return false;
        }

        console.log(`returned: ${sheet} from ${affectedWorkbook}`);

        return resultSheet;
    }

    // return active sheet name
    getActiveSheetName() {

        // no sheet selected
        if (!this.activeSheet) {
            console.error("no sheet selected");

            return false;
        }

        console.log("active sheet:", this.activeSheet);

        return true;
    }

    /* fetching data */

    // fetch data from cell
    fetch(ref, format = "value") {

        // no sheet selected
        if (!this.activeSheet) {
            console.log("no active sheet selected");

            return false;
        }

        let col, row;

        // "A1" format
        if (typeof ref === "string") {
            [[col, row]] = this.#refToAddress(ref);
        }
        // [1, 2] format
        else if (Array.isArray(ref) && ref.length === 2) {
            col = ref[0];
            row = ref[1];
        }
        // invalid format
        else {
            console.log("invalid reference format:", ref);
        }

        const sheet = this.workbooks[this.activeWorkbook].getWorksheet(this.activeSheet);
        const cell = sheet.getCell(row, col);

        // return cell, value or formula
        if (format === "cell") {
            return cell;
        } else {
            const valueType = typeof cell.value;

            // return cell value
            if (format == "value") {
                return valueType !== "object" ? cell.value : cell.value?.result;
            }
            // return formula if available
            else if (format === "formula") {
                return valueType !== "object" ? cell.value : cell.value?.formula;
            } 
            // return cell styles
            else if (format === "style") {
                return cell.style;
            }
            // invalid format
            else {
                return false;
            }
        }
    }

    // fetch data from cell range
    fetchRange(refs, format = "value") {
        const range = [];

        // ensure refs is loopable
        if (!Array.isArray(refs)) {
            refs = [refs];
        } 
        // [1, 2] format
        else if (refs.length === 2) {
            range.push(this.fetch([refs[0], refs[1]], format));
            return range;
        }

        // loop over references
        for (const ref of refs) {

            // ["A1:B2"] format
            if (typeof ref === "string" && ref.includes(":")) {
                const [startRef, endRef] = ref.split(":");

                // get bounds addresses
                let [
                    [startCol, startRow],
                    [endCol, endRow]
                ] = this.#refToAddress([startRef, endRef]);

                // flip start & end columns if reversed
                if (startCol > endCol) {
                    const newCol = startCol;
                    startCol = endCol;
                    endCol = newCol;
                }

                // flip start & end rows if reversed
                if (startRow > endRow) {
                    const newRow = startRow;
                    startRow = endRow;
                    endRow = newRow;
                }

                // loop over rows
                for (let row = startRow; row <= endRow; row++) {
                    const rowRange = [];

                    // loop over columns
                    for (let col = startCol; col <= endCol; col++) {
                        const cell = this.fetch([col, row], format);
                        rowRange.push(cell);
                    }

                    range.push(rowRange);
                }
            }
            // ["A1"] or [[1, 2], [3, 4]] format
            else {
                range.push(this.fetch(ref, format));
            }
        }

        return range;
    }

    /* setting data */

    set() { }
    setRange() { }

    /* clearing data */

    clear() { }

    /* copying data */

    copy() { }
    copyRange() { }

    /* utilities */

    fillValue() { }
    fillFormula() { }

    /* styling */

    setStyle() { }
    setStyles() { }

    /* private */

    #refToAddress(refs) {
        const addresses = [];

        // ensure refs is loopable
        if (!Array.isArray(refs)) refs = [refs];

        // loop over references
        for (const ref of refs) {

            // reference is not a string
            if (!typeof ref === "string") {
                console.error("reference of invalid type:", typeof ref);

                return [];
            }

            // split string according to template, extract columns & rows 
            const match = /^([A-Z]+)(\d+)$/.exec(ref.toUpperCase());

            // check if regex valid
            if (!match) {
                console.error("regex failed with string:", ref);

                return [];
            }

            // deconstruct regex output
            const [, col, row] = match;

            let colAddress = 0;
            let rowAddress = parseInt(row, 10);

            // loop over column reference letters
            for (let i = 0; i < col.length; i++) {

                // convert to base-26
                colAddress *= 26;

                // A = 65 in ASCII so subtract 64 to start at column 1 for A
                colAddress += col.charCodeAt(i) - 64;
            }

            // add address to output array
            addresses.push([colAddress, rowAddress]);
        }

        return addresses;
    }
}