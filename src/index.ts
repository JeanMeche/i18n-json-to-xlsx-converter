#!/usr/bin/env node
import * as Excel from 'exceljs';
const path = require('path');
const fs = require('fs-extra');
const chalk = require('chalk');
const readXLSXFile = require('read-excel-file/node');
const unflatten = require('flat').unflatten;
import utils from './utils';

(async () => {
  try {
    const basepath = process.argv.slice(3);
    const filePath = utils.createPathByCheckingSpaceCharacter(basepath);

    if (!filePath || typeof filePath === 'boolean') {
      utils.parseErrorMessage('No file path to convert is given. Specify the file path after the --convert parameter.');
      process.exit(1);
    }

    const sourceFileType = utils.getSourceFileType(filePath);
    const isMultipleJSONFilePaths = utils.getJSONFilePaths(filePath).length > 1;
    const isMultipleJSONFilePathsValid = utils.isMultipleJSONFilePathsValid(filePath);

    if (utils.isJSON(sourceFileType) || utils.isXLSX(sourceFileType) || isMultipleJSONFilePathsValid) {
      utils.createProcessMessageByType(filePath, sourceFileType, isMultipleJSONFilePathsValid && isMultipleJSONFilePaths);
    } else {
      utils.checkForMultipleJSONFileErrors(filePath, process);

      utils.parseErrorMessage('File type is not supported. Either use JSON or XLSX file to convert.');
      process.exit(1);
    }

    if (utils.isXLSX(sourceFileType)) {
      const readXlsx = () => {
        return readXLSXFile(filePath).then((rows: string[][]) => {
          const titleRow = rows[0];
          const allLanguages: any = {};
          const titles = [];

          for (const [idx, row] of titleRow.entries()) {
            titles.push(row);

            if (idx > 0) {
              allLanguages[row] = {};
            }
          }

          for (let idx = 1; idx < rows.length; idx++) {
            const row = rows[idx];

            for (let secondIdx = 1; secondIdx < row.length; secondIdx++) {
              if (row[0]) {
                allLanguages[titles[secondIdx]][row[0]] = row[secondIdx];
              }
            }
          }

          return allLanguages;
        });
      };

      readXlsx()
        .then((allLanguages: any) => {
          let outputFileName = '';

          for (const languageTitle in allLanguages) {
            outputFileName = `${languageTitle.trim().toLowerCase()}${utils.getFileExtension(filePath)}`;

            const unflattenedLanguageObj = unflatten(allLanguages[languageTitle], { object: true });

            fs.writeFileSync(utils.documentSavePath(filePath, outputFileName), JSON.stringify(unflattenedLanguageObj, null, 2), 'utf-8');

            utils.log(chalk.yellow(`Output file name for ${languageTitle} is ${outputFileName}`));
            utils.log(chalk.grey(`Location of the created file is`));
            utils.log(chalk.magentaBright(`${utils.documentSavePath(filePath, outputFileName)}\n`));
          }

          utils.log(chalk.green(`File conversion is successful!`));
        })
        .catch((e: Error) => {
          utils.error(chalk.red(`Error: ${e}`));

          process.exit(1);
        });
    } else {
      const JSONFiles = utils.getJSONFilePaths(filePath);
      const workbook = new Excel.Workbook();
      const worksheet = workbook.addWorksheet();
      const rowForKey = new Map<string, number>();

      for (const JSONFile of JSONFiles!) {
        const colIndex = JSONFiles.indexOf(JSONFile) + 2; // starts at 1, skipping key column
        const lang = utils.getFileName(JSONFile).toUpperCase();
        const sourceBuffer = await fs.promises.readFile(JSONFile);
        const sourceText = sourceBuffer.toString();
        const sourceData = JSON.parse(sourceText);

        const writeToXLSX = (key: string, value: string) => {
          let rowIndex: number;
          if (rowForKey.has(key)) {
            rowIndex = rowForKey.get(key)!;
          } else {
            rowIndex = rowForKey.size;
            rowForKey.set(key, rowIndex);
          }

          const rows = worksheet.getRow(rowIndex);
          rows.getCell(1).value = key;
          // Check for null, "" of the values and assign semantic character for that
          rows.getCell(colIndex).value = (value || '-').toString();
        };

        writeToXLSX('Key', lang);

        const parseAndWrite = (parentKey: string | null, targetObject: any) => {
          const keys = Object.keys(targetObject);

          for (const key of keys as string[]) {
            const element: any = targetObject[key];

            if (typeof element === 'object' && element !== null) {
              parseAndWrite(utils.writeByCheckingParent(parentKey, key), element);
            } else {
              writeToXLSX(utils.writeByCheckingParent(parentKey, key), element);
            }
          }
        };

        parseAndWrite(null, sourceData);

        worksheet.getColumn(colIndex).width = 50;
      }
      worksheet.getColumn(1).width = 50;

      const JSONFile = JSONFiles[0];
      const outputFilename = `translations.xlsx`;
      await workbook.xlsx
        .writeFile(utils.documentSavePath(JSONFile, outputFilename))
        .then(() => {
          utils.log(chalk.yellow(`Output file name is ${outputFilename}`));
          utils.log(chalk.grey(`Location of the created file is`));
          utils.log(chalk.magentaBright(`${utils.documentSavePath(JSONFile, `${outputFilename}`)}\n`));
          utils.log(chalk.green(`File conversion is successful!\n`));
        })
        .catch((e: Error) => {
          utils.error(chalk.red(`Error: ${e}`));

          process.exit(1);
        });
    }
  } catch (e) {
    utils.error(chalk.red(`Error: ${e}`));

    process.exit(1);
  }
})();
