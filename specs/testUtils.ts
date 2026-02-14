import { dirname, join } from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));

const outputDir = join(__dirname, 'output');
const writeTestFile = join(outputDir, 'writetest.xls');
const tempFile = join(outputDir, 'tempfile');
const filesDir = join(__dirname, 'files');
const testPicture = join(filesDir, 'dummy.jpg');
const xlsxTableFile = join(filesDir, 'table.xlsx');
const xlsmFormControlFile = join(filesDir, 'control.xlsm');

export const initFilesystem = () => {
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir);
    }

    if (fs.existsSync(writeTestFile)) {
        fs.unlinkSync(writeTestFile);
    }
};

export const getWriteTestFile = () => writeTestFile;

export const getTempFile = () => tempFile;

export const getXlsxTableFile = () => xlsxTableFile;

export const getXlsmFormControlFile = () => xlsmFormControlFile;

export const getTestPicturePath = () => testPicture;

export const compareBuffers = (buf1: Buffer, buf2: Buffer): boolean => {
    if (buf1.length !== buf2.length) return false;

    for (let i = 0; i < buf1.length; i++) {
        if (buf1[i] !== buf2[i]) return false;
    }

    return true;
};

export const testPictureWidth = 640;
export const testPictureHeight = 480;

export const epsilon = 1e-9;
