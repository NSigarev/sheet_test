import { Injectable } from '@nestjs/common';
import axios from 'axios';
import { google } from 'googleapis';
interface ApiRes {
  ObjectName: string;
  Year: number;
  Month: number;
  Plan: number;
  Fact: number;
}

@Injectable()
export class AppService {
  async getData(): Promise<string> {
    try {
      const apiUrl =
        'https://script.google.com/macros/s/AKfycbxBLv4QkfvO18qyD52fmkt34tJ29YTp1aMUifIwEHVzmiYZciEazIVKI1Q0VkA5Jiu9/exec?apikey=640fb8c3-cc56-4ade-a447-8313d65657ee&action=get_data';

      const response = await axios.get<ApiRes[]>(apiUrl);
      const data = response.data;
      const vadata = data.filter((item) => item.ObjectName.includes('ВА'));
      const bdata = data.filter((item) => item.ObjectName.includes('Б'));

      const auth = new google.auth.JWT(
        process.env.GOOGLE_CLIENT_EMAIL,
        undefined,
        (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n'),
        [
          'https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive',
        ],
      );

      const sheets = google.sheets({ version: 'v4', auth });
      const drive = google.drive({ version: 'v3', auth });

      const sourceSpreadsheetId =
        '1gdQKfy6r2zqvym6PFJJEQ8entjTwdMthelEo2sSnJ_M';
      const copyResponse = await drive.files.copy({
        fileId: sourceSpreadsheetId,
        requestBody: {
          name: `Data Copy ${new Date().toISOString()}`,
        },
      });

      const copiedSpreadsheetId = copyResponse.data.id;
      if (!copiedSpreadsheetId) {
        throw new Error('Ошибка при копировании таблицы');
      }

      await drive.permissions.create({
        fileId: copiedSpreadsheetId,
        requestBody: {
          role: 'reader',
          type: 'anyone',
        },
      });

      const sheetData = [
        { sheetName: 'ВА', data: vadata || [] },
        { sheetName: 'Б', data: bdata || [] },
      ];

      for (const { sheetName, data: entries } of sheetData) {
        // Получаем значения из столбца B (где ObjectName)
        const sheetRes = await sheets.spreadsheets.values.get({
          spreadsheetId: copiedSpreadsheetId,
          range: `${sheetName}!B1:B1000`,
        });

        const bColumn = sheetRes.data.values || [];

        const rowsMap: Map<number, (string | number)[]> = new Map();

        for (const entry of entries) {
          const rowIndex = bColumn.findIndex(
            (row) => row[0] === entry.ObjectName,
          );
          if (rowIndex === -1) {
            console.warn(
              `ObjectName "${entry.ObjectName}" не найден на листе "${sheetName}"`,
            );
            continue;
          }

          const targetRow = rowIndex + 1;
          const yearOffset = (entry.Year - 2024) * 12 * 2;
          const monthOffset = (entry.Month - 1) * 2;

          const planCol = 6 + yearOffset + monthOffset;
          const factCol = planCol + 1;

          const row = rowsMap.get(targetRow) || [];
          row[planCol] = entry.Plan;
          row[factCol] = entry.Fact;
          rowsMap.set(targetRow, row);
        }

        const updates = Array.from(rowsMap.entries()).map(
          ([rowNumber, rowValues]) => {
            const colToLetter = (col: number): string => {
              let result = '';
              while (col >= 0) {
                result = String.fromCharCode((col % 26) + 65) + result;
                col = Math.floor(col / 26) - 1;
              }
              return result;
            };

            const maxColIndex = rowValues.length - 1;
            const range = `${sheetName}!G${rowNumber}:${colToLetter(maxColIndex)}${rowNumber}`;

            return {
              range,
              values: [rowValues.slice(6)],
            };
          },
        );

        if (updates.length > 0) {
          await sheets.spreadsheets.values.batchUpdate({
            spreadsheetId: copiedSpreadsheetId,
            requestBody: {
              valueInputOption: 'RAW',
              data: updates,
            },
          });
        }
      }

      return `https://docs.google.com/spreadsheets/d/${copiedSpreadsheetId}`;
    } catch (err: any) {
      console.error('Ошибка при получении или записи данных:', err.message);
      throw new Error('Ошибка сервиса getData');
    }
  }
}
