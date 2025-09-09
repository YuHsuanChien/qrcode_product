import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";

interface ExcelData {
	success: boolean;
	fileName: string;
	filePath: string;
	rowCount: number;
	columnCount: number;
	data: any[];
	readAt: string;
	worksheetName?: string;
	totalWorksheets?: number;
	worksheetNames?: string[];
	error?: Error;
}

interface ReadOptions {
	worksheetIndex?: number; // 讀取第幾個工作表 (從 1 開始)
	worksheetName?: string; // 或指定工作表名稱
	includeEmpty?: boolean; // 是否包含空白儲存格
	header?: boolean; // 是否將第一行當作標題
	raw?: boolean; // 是否保留原始格式 (日期、數字等)
}

export default class ReadExcelFile {
	checkFileExists(filePath: string): boolean {
		try {
			return fs.existsSync(filePath);
		} catch (error) {
			console.error("檢查檔案存在時發生錯誤：", error);
			return false;
		}
	}

	/**
	 * 讀取單一 Excel 檔案
	 * @param {string} filePath - Excel 檔案路徑
	 * @param {ReadOptions} options - 讀取選項
	 * @returns {Promise<ExcelData>} 讀取結果
	 */
	async readSingleExcelFile(
		filePath: string,
		options: ReadOptions = {}
	): Promise<ExcelData> {
		try {
			console.log(`📖 正在讀取：${filePath}`);

			if (!this.checkFileExists(filePath)) {
				throw new Error(`檔案不存在：${filePath}`);
			}

			// 檢查檔案擴展名
			const ext = path.extname(filePath).toLowerCase();
			if (![".xlsx", ".xls"].includes(ext)) {
				throw new Error(`不支援的檔案格式：${ext}，僅支援 .xlsx, .xls`);
			}

			const workbook = new ExcelJS.Workbook();
			await workbook.xlsx.readFile(filePath);

			// 取得所有工作表名稱
			const worksheetNames = workbook.worksheets.map((ws) => ws.name);

			// 決定要讀取哪個工作表
			let worksheet: ExcelJS.Worksheet;
			let worksheetName: string;

			if (options.worksheetName) {
				const targetWorksheet = workbook.getWorksheet(options.worksheetName);
				if (!targetWorksheet) {
					throw new Error(`找不到工作表：${options.worksheetName}`);
				}
				worksheet = targetWorksheet; // 這裡型別已經是 Worksheet
				worksheetName = options.worksheetName;
			} else {
				const index = options.worksheetIndex || 1;
				const targetWorksheet = workbook.getWorksheet(index);
				if (!targetWorksheet) {
					throw new Error(`找不到第 ${index} 個工作表`);
				}
				worksheet = targetWorksheet; // 這裡型別已經是 Worksheet
				worksheetName = targetWorksheet.name;
			}

			// 轉換資料
			const data = this.worksheetToArray(worksheet, options);

			console.log(
				`✅ 成功讀取：${path.basename(filePath)} - 工作表：${worksheetName} (${
					data.length
				} 行)`
			);

			return {
				success: true,
				fileName: path.basename(filePath),
				filePath: filePath,
				rowCount: data.length,
				columnCount: data.length > 0 ? data[0].length : 0,
				data: data,
				worksheetName: worksheetName,
				totalWorksheets: workbook.worksheets.length,
				worksheetNames: worksheetNames,
				readAt: new Date().toISOString(),
			};
		} catch (error) {
			console.error(`❌ 讀取失敗：${path.basename(filePath)} - ${error}`);

			return {
				success: false,
				fileName: path.basename(filePath),
				filePath: filePath,
				rowCount: 0,
				columnCount: 0,
				data: [],
				error: error as Error,
				readAt: new Date().toISOString(),
			};
		}
	}

	/**
	 * 讀取所有工作表
	 * @param {string} filePath - Excel 檔案路徑
	 * @param {ReadOptions} options - 讀取選項
	 * @returns {Promise<ExcelData[]>} 所有工作表的讀取結果
	 */
	async readAllWorksheets(
		filePath: string,
		options: ReadOptions = {}
	): Promise<ExcelData[]> {
		try {
			if (!this.checkFileExists(filePath)) {
				throw new Error(`檔案不存在：${filePath}`);
			}

			const workbook = new ExcelJS.Workbook();
			await workbook.xlsx.readFile(filePath);

			const results: ExcelData[] = [];

			for (const worksheet of workbook.worksheets) {
				try {
					const data = this.worksheetToArray(worksheet, options);

					results.push({
						success: true,
						fileName: path.basename(filePath),
						filePath: filePath,
						rowCount: data.length,
						columnCount: data.length > 0 ? data[0].length : 0,
						data: data,
						worksheetName: worksheet.name,
						totalWorksheets: workbook.worksheets.length,
						worksheetNames: workbook.worksheets.map((ws) => ws.name),
						readAt: new Date().toISOString(),
					});
				} catch (error) {
					results.push({
						success: false,
						fileName: path.basename(filePath),
						filePath: filePath,
						rowCount: 0,
						columnCount: 0,
						data: [],
						worksheetName: worksheet.name,
						error: error as Error,
						readAt: new Date().toISOString(),
					});
				}
			}

			return results;
		} catch (error) {
			return [
				{
					success: false,
					fileName: path.basename(filePath),
					filePath: filePath,
					rowCount: 0,
					columnCount: 0,
					data: [],
					error: error as Error,
					readAt: new Date().toISOString(),
				},
			];
		}
	}

	/**
	 * 將工作表轉換為陣列
	 * @private
	 */
	private worksheetToArray(
		worksheet: ExcelJS.Worksheet,
		options: ReadOptions
	): any[][] {
		const data: any[][] = [];
		const { includeEmpty = true, raw = false } = options;

		worksheet.eachRow((row, rowNumber) => {
			const rowData: any[] = [];

			// 取得這一行的最大列數
			const maxCol = Math.max(row.cellCount, worksheet.columnCount || 0);

			for (let colNumber = 1; colNumber <= maxCol; colNumber++) {
				const cell = row.getCell(colNumber);
				let value: any;

				if (raw) {
					// 保留原始值和格式
					value = cell.value;
				} else {
					// 轉換為簡單值
					if (cell.value === null || cell.value === undefined) {
						value = includeEmpty ? null : "";
					} else if (typeof cell.value === "object" && "result" in cell.value) {
						// 處理公式
						value = cell.value.result;
					} else if (cell.value instanceof Date) {
						// 處理日期
						value = cell.value.toISOString();
					} else {
						value = cell.value;
					}
				}

				rowData.push(value);
			}

			data.push(rowData);
		});

		return data;
	}

	/**
	 * 取得工作表資訊（不讀取資料）
	 * @param {string} filePath - Excel 檔案路徑
	 * @returns {Promise<object>} 工作表資訊
	 */
	async getWorksheetInfo(filePath: string) {
		try {
			if (!this.checkFileExists(filePath)) {
				throw new Error(`檔案不存在：${filePath}`);
			}

			const workbook = new ExcelJS.Workbook();
			await workbook.xlsx.readFile(filePath);

			const worksheets = workbook.worksheets.map((ws) => ({
				name: ws.name,
				id: ws.id,
				rowCount: ws.rowCount,
				columnCount: ws.columnCount,
				actualRowCount: ws.actualRowCount,
				actualColumnCount: ws.actualColumnCount,
			}));

			return {
				success: true,
				fileName: path.basename(filePath),
				totalWorksheets: workbook.worksheets.length,
				worksheets: worksheets,
			};
		} catch (error) {
			return {
				success: false,
				fileName: path.basename(filePath),
				error: error as Error,
			};
		}
	}
}
