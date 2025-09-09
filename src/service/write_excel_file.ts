import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";

interface WriteOptions {
	worksheetName?: string;
	overwrite?: boolean;
	autoFilter?: boolean;
	freezeHeader?: boolean;
}

interface ImageOptions {
	imagePath: string;
	cell: string;
	width?: number;
	height?: number;
	maintainAspectRatio?: boolean;
}

interface WriteResult {
	success: boolean;
	fileName: string;
	filePath: string;
	rowCount?: number;
	columnCount?: number;
	worksheetName?: string;
	imagesInserted?: number;
	createdAt?: string;
	updatedAt?: string;
	error?: Error;
}

export default class WriteExcelFile {
	async clearFolder(folderPath: string) {
		try {
			if (fs.existsSync(folderPath)) {
				const files = fs.readdirSync(folderPath);
				files.forEach((file) => {
					const filePath = path.join(folderPath, file);
					if (fs.statSync(filePath).isFile()) {
						fs.unlinkSync(filePath);
						console.log(`🗑️ 已刪除檔案：${filePath}`);
					}
				});
			}
		} catch (error) {
			console.error("刪除檔案存在時發生錯誤：", error);
			return false;
		}
	}

	checkFileExists(filePath: string): boolean {
		try {
			return fs.existsSync(filePath);
		} catch (error) {
			console.error("檢查檔案存在時發生錯誤：", error);
			return false;
		}
	}

	/**
	 * 寫入 Excel 檔案
	 */
	async writeExcelFile(
		filePath: string,
		data: any[][],
		options: WriteOptions = {}
	): Promise<WriteResult> {
		try {
			const {
				worksheetName = "Sheet1",
				overwrite = true,
				autoFilter = false,
				freezeHeader = false,
			} = options;

			if (!overwrite && this.checkFileExists(filePath)) {
				throw new Error(
					`檔案已存在：${filePath}，請設定 overwrite: true 來覆寫`
				);
			}

			console.log(`📝 正在寫入：${filePath}`);

			const workbook = new ExcelJS.Workbook();
			const worksheet = workbook.addWorksheet(worksheetName);

			data.forEach((row, index) => {
				worksheet.addRow(row);
				if (index === 0 && (autoFilter || freezeHeader)) {
					const headerRow = worksheet.getRow(1);
					headerRow.font = { bold: true };
				}
			});

			if (autoFilter && data.length > 0) {
				worksheet.autoFilter = {
					from: "A1",
					to: { row: data.length, column: data[0].length },
				};
			}

			if (freezeHeader && data.length > 0) {
				worksheet.views = [{ state: "frozen", ySplit: 1 }];
			}

			worksheet.columns.forEach((column, index) => {
				let maxLength = 10;
				data.forEach((row) => {
					if (row[index] && row[index].toString().length > maxLength) {
						maxLength = Math.min(row[index].toString().length + 2, 50);
					}
				});
				if (column) {
					column.width = maxLength;
				}
			});

			await workbook.xlsx.writeFile(filePath);

			console.log(
				`✅ 成功寫入：${path.basename(filePath)} (${data.length} 行)`
			);

			return {
				success: true,
				fileName: path.basename(filePath),
				filePath: filePath,
				rowCount: data.length,
				columnCount: data.length > 0 ? data[0].length : 0,
				worksheetName: worksheetName,
				createdAt: new Date().toISOString(),
			};
		} catch (error) {
			console.error(`❌ 寫入失敗：${path.basename(filePath)} - ${error}`);
			return {
				success: false,
				fileName: path.basename(filePath),
				filePath: filePath,
				error: error as Error,
				createdAt: new Date().toISOString(),
			};
		}
	}

	/**
	 * 解析儲存格地址 (如: 'A1' -> {col: 1, row: 1})
	 */
	private parseCellAddress(cell: string): { col: number; row: number } {
		const match = cell.match(/^([A-Z]+)(\d+)$/);
		if (!match) {
			throw new Error(`無效的儲存格地址：${cell}`);
		}

		const colStr = match[1];
		const rowStr = match[2];

		let col = 0;
		for (let i = 0; i < colStr.length; i++) {
			col = col * 26 + (colStr.charCodeAt(i) - "A".charCodeAt(0) + 1);
		}

		const row = parseInt(rowStr, 10);
		return { col, row };
	}

	/**
	 * 檢查圖片檔案是否有效
	 */
	private isValidImageFile(imagePath: string): boolean {
		try {
			if (!fs.existsSync(imagePath)) {
				console.warn(`圖片不存在: ${imagePath}`);
				return false;
			}

			const stats = fs.statSync(imagePath);
			if (stats.size === 0) {
				console.warn(`圖片檔案為空: ${imagePath}`);
				return false;
			}

			// 檢查檔案大小限制 (10MB)
			if (stats.size > 10 * 1024 * 1024) {
				console.warn(`圖片檔案過大: ${imagePath} (${stats.size} bytes)`);
				return false;
			}

			const ext = path.extname(imagePath).toLowerCase();
			if (![".png", ".jpg", ".jpeg", ".gif", ".bmp"].includes(ext)) {
				console.warn(`不支援的圖片格式: ${ext}`);
				return false;
			}

			return true;
		} catch (error) {
			console.warn(`檢查圖片檔案時發生錯誤: ${imagePath} - ${error}`);
			return false;
		}
	}

	/**
	 * 安全地插入圖片到 Excel 檔案
	 */
	async insertImagesSafely(
		excelPath: string,
		images: ImageOptions[],
		worksheetName?: string
	): Promise<WriteResult> {
		let backupPath: string | null = null;

		try {
			console.log(`🖼️ 開始安全插入 ${images.length} 張圖片到: ${excelPath}`);

			if (!this.checkFileExists(excelPath)) {
				throw new Error(`Excel 檔案不存在: ${excelPath}`);
			}

			// 1. 過濾有效圖片
			const validImages = images.filter((img) =>
				this.isValidImageFile(img.imagePath)
			);
			console.log(`📊 有效圖片數量: ${validImages.length}/${images.length}`);

			if (validImages.length === 0) {
				return {
					success: true,
					fileName: path.basename(excelPath),
					filePath: excelPath,
					imagesInserted: 0,
					updatedAt: new Date().toISOString(),
				};
			}

			// 2. 載入工作簿
			const workbook = new ExcelJS.Workbook();
			await workbook.xlsx.readFile(excelPath);

			// 3. 選擇工作表
			const worksheet = worksheetName
				? workbook.getWorksheet(worksheetName)
				: workbook.getWorksheet(1);

			if (!worksheet) {
				throw new Error(`找不到工作表: ${worksheetName || "第一個工作表"}`);
			}

			console.log(`📋 使用工作表: ${worksheet.name}`);

			// 4. 批次插入圖片
			let insertedCount = 0;
			const batchSize = 5; // 每批處理5張圖片

			for (let i = 0; i < validImages.length; i += batchSize) {
				const batch = validImages.slice(
					i,
					Math.min(i + batchSize, validImages.length)
				);

				console.log(
					`🔄 處理批次 ${Math.floor(i / batchSize) + 1}/${Math.ceil(
						validImages.length / batchSize
					)}`
				);

				for (const imageOption of batch) {
					try {
						const success = await this.insertSingleImage(
							workbook,
							worksheet,
							imageOption
						);
						if (success) {
							insertedCount++;
						}
					} catch (imgError) {
						console.error(
							`插入圖片失敗: ${imageOption.imagePath} - ${imgError}`
						);
					}
				}

				// 批次間稍作休息
				if (i + batchSize < validImages.length) {
					await new Promise((resolve) => setTimeout(resolve, 300));
				}
			}

			// 5. 儲存檔案
			console.log(`💾 正在儲存修改後的檔案...`);
			await workbook.xlsx.writeFile(excelPath);

			console.log(
				`🎉 圖片插入完成: ${insertedCount}/${validImages.length} 張成功`
			);

			return {
				success: true,
				fileName: path.basename(excelPath),
				filePath: excelPath,
				imagesInserted: insertedCount,
				worksheetName: worksheet.name,
				updatedAt: new Date().toISOString(),
			};
		} catch (error) {
			console.error(`❌ 圖片插入失敗: ${error}`);

			// 嘗試恢復備份
			if (backupPath && this.checkFileExists(backupPath)) {
				try {
					fs.copyFileSync(backupPath, excelPath);
					console.log(`🔄 已從備份恢復原始檔案`);
				} catch (restoreError) {
					console.error(`恢復備份失敗: ${restoreError}`);
				}
			}

			return {
				success: false,
				fileName: path.basename(excelPath),
				filePath: excelPath,
				error: error as Error,
				updatedAt: new Date().toISOString(),
			};
		}
	}

	/**
	 * 插入單張圖片
	 */
	private async insertSingleImage(
		workbook: ExcelJS.Workbook,
		worksheet: ExcelJS.Worksheet,
		imageOption: ImageOptions
	): Promise<boolean> {
		try {
			const {
				imagePath,
				cell,
				width = 50, // 較小的預設尺寸
				height = 50,
				maintainAspectRatio = true,
			} = imageOption;

			// 解析儲存格位置
			const cellInfo = this.parseCellAddress(cell);

			// 讀取圖片檔案
			const imageBase64 = fs.readFileSync(imagePath, { encoding: "base64" });
			const ext = path.extname(imagePath).toLowerCase().replace(".", "");

			// 標準化副檔名
			let standardExt: "jpeg" | "png" | "gif";
			switch (ext) {
				case "jpg":
				case "jpeg":
					standardExt = "jpeg";
					break;
				case "png":
					standardExt = "png";
					break;
				case "gif":
					standardExt = "gif";
					break;
				default:
					throw new Error(`不支援的圖片格式: ${ext}`);
			}

			// 加入圖片到工作簿
			const imageId = workbook.addImage({
				base64: imageBase64,
				extension: standardExt,
			});

			// 設置圖片位置和大小
			const imageConfig = {
				tl: {
					col: cellInfo.col - 1,
					row: cellInfo.row - 1,
				},
				ext: {
					width: maintainAspectRatio ? width : width,
					height: maintainAspectRatio ? width : height, // 保持比例時使用相同值
				},
			};

			// 插入圖片
			worksheet.addImage(imageId, imageConfig);

			console.log(`✅ 圖片插入成功: ${path.basename(imagePath)} → ${cell}`);
			return true;
		} catch (error) {
			console.error(`插入圖片失敗: ${imageOption.imagePath} - ${error}`);
			return false;
		}
	}

	/**
	 * 驗證 Excel 檔案完整性
	 */
	async validateExcelFile(filePath: string): Promise<boolean> {
		try {
			console.log(`🔍 驗證 Excel 檔案: ${filePath}`);

			const workbook = new ExcelJS.Workbook();
			await workbook.xlsx.readFile(filePath);

			console.log(`✅ Excel 檔案驗證通過`);
			return true;
		} catch (error) {
			console.error(`❌ Excel 檔案驗證失敗: ${error}`);
			return false;
		}
	}

	/**
	 * 舊版插入方法 (保持向下相容)
	 */
	async insertImages(
		excelPath: string,
		images: ImageOptions[],
		worksheetName?: string
	): Promise<WriteResult> {
		// 直接調用安全版本
		return this.insertImagesSafely(excelPath, images, worksheetName);
	}
}
