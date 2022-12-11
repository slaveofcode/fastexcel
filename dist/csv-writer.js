const fs = require('fs');

class CsvWriter {
    constructor(filePath, headers) {
        this.file = fs.createWriteStream(filePath);
        let row = '';
        for (const col of headers) {
            if (row !== '') {
                row += ','
            }
            row += col
        }
        row += '\n';
        this.file.write(row);
    }

    async write(row) {
        let rowLine = '';
        for (const col of row) {
            if (rowLine !== '') {
                rowLine += ','
            }
            rowLine += col;
        }
        rowLine += '\n';
        await new Promise((res, rej) => {
            if (this.file.write(rowLine)) {
                process.nextTick(res);
            } else {
                this.file.once('drain', () => {
                    this.file.off('error', rej);
                    res(undefined);
                })
                this.file.once('error', rej);
            }
        })
    }

    pack() {
        this.file.end();
        return new Promise((res, rej) => {
            this.file.on("finish", res);
            this.file.on("error", rej);
        });
    }
}

module.exports = { CsvWriter };
