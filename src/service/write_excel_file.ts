import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";

interface WriteOptions {
	worksheetName?: string; // 工作表名稱
	overwrite?: boolean; // 是否覆寫現有檔案
	autoFilter?: boolean; // 是否添加自動篩選
	freezeHeader?: boolean; // 是否凍結標題行
}

interface ImageOptions {
	imagePath: string; // 圖片檔案路徑
	cell: string; // 要插入的儲存格 (如: 'A1')
	width?: number; // 圖片寬度 (像素)
	height?: number; // 圖片高度 (像素)
	maintainAspectRatio?: boolean; // 保持長寬比
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
	 * @param {string} filePath - 輸出檔案路徑
	 * @param {any[][]} data - 要寫入的資料
	 * @param {WriteOptions} options - 寫入選項
	 * @returns {Promise<WriteResult>} 寫入結果
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

			// 檢查檔案是否存在
			if (!overwrite && this.checkFileExists(filePath)) {
				throw new Error(
					`檔案已存在：${filePath}，請設定 overwrite: true 來覆寫`
				);
			}

			console.log(`📝 正在寫入：${filePath}`);

			const workbook = new ExcelJS.Workbook();
			const worksheet = workbook.addWorksheet(worksheetName);

			// 寫入資料
			data.forEach((row, index) => {
				worksheet.addRow(row);

				// 如果是第一行且設定為標題，加粗
				if (index === 0 && (autoFilter || freezeHeader)) {
					const headerRow = worksheet.getRow(1);
					headerRow.font = { bold: true };
				}
			});

			// 添加自動篩選
			if (autoFilter && data.length > 0) {
				worksheet.autoFilter = {
					from: "A1",
					to: { row: data.length, column: data[0].length },
				};
			}

			// 凍結標題行
			if (freezeHeader && data.length > 0) {
				worksheet.views = [{ state: "frozen", ySplit: 1 }];
			}

			// 自動調整欄寬
			worksheet.columns.forEach((column, index) => {
				let maxLength = 10; // 最小寬度
				data.forEach((row) => {
					if (row[index] && row[index].toString().length > maxLength) {
						maxLength = Math.min(row[index].toString().length + 2, 50); // 最大寬度 50
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
	 * 在現有Excel檔案中插入圖片
	 * @param {string} excelPath - Excel 檔案路徑
	 * @param {ImageOptions[]} images - 圖片選項陣列
	 * @param {string} worksheetName - 工作表名稱 (可選)
	 * @returns {Promise<WriteResult>} 插入結果
	 */
	async insertImages(
		excelPath: string,
		images: ImageOptions[],
		worksheetName?: string
	): Promise<WriteResult> {
		try {
			console.log(`🖼️ 正在插入 ${images.length} 張圖片到：${excelPath}`);

			if (!this.checkFileExists(excelPath)) {
				throw new Error(`Excel 檔案不存在：${excelPath}`);
			}

			// 檢查所有圖片檔案是否存在
			const validImages: ImageOptions[] = [];
			for (const img of images) {
				if (this.checkFileExists(img.imagePath)) {
					validImages.push(img);
				} else {
					console.warn(`⚠️ 圖片檔案不存在，跳過：${img.imagePath}`);
				}
			}

			if (validImages.length === 0) {
				console.warn("⚠️ 沒有有效的圖片檔案可插入");
				return {
					success: true,
					fileName: path.basename(excelPath),
					filePath: excelPath,
					imagesInserted: 0,
					updatedAt: new Date().toISOString(),
				};
			}

			const workbook = new ExcelJS.Workbook();
			await workbook.xlsx.readFile(excelPath);

			// 選擇工作表
			const worksheet = worksheetName
				? workbook.getWorksheet(worksheetName)
				: workbook.getWorksheet(1);

			if (!worksheet) {
				throw new Error(`找不到工作表：${worksheetName || "第一個工作表"}`);
			}

			// 插入每張圖片
			let insertedCount = 0;
			for (const imageOption of validImages) {
				try {
					const {
						imagePath,
						cell,
						width = 80,
						height = 80,
						maintainAspectRatio = true,
					} = imageOption;

					// 取得圖片副檔名
					const ext = path.extname(imagePath).toLowerCase().replace(".", "");

					if (!["png", "jpg", "jpeg", "gif", "bmp"].includes(ext)) {
						console.warn(`⚠️ 不支援的圖片格式：${ext}，跳過：${imagePath}`);
						continue;
					}

					// 加入圖片到工作簿
					const imageId = workbook.addImage({
						filename: imagePath,
						extension: "png",
					});

					// 解析儲存格位置
					const cellInfo = this.parseCellAddress(cell);

					// 設定圖片位置和大小
					const imageConfig: any = {
						tl: {
							col: cellInfo.col - 1,
							row: cellInfo.row - 1,
						},
						ext: { width, height },
					};

					// 如果要保持長寬比，只設定寬度
					if (maintainAspectRatio) {
						delete imageConfig.ext.height;
					}

					worksheet.addImage(imageId, imageConfig);
					insertedCount++;

					console.log(
						`✅ 圖片插入成功：${path.basename(imagePath)} → 儲存格 ${cell}`
					);
				} catch (imgError) {
					console.error(
						`❌ 插入圖片失敗：${imageOption.imagePath} - ${imgError}`
					);
				}
			}

			// 儲存檔案
			await workbook.xlsx.writeFile(excelPath);

			console.log(
				`🎉 圖片插入完成：${insertedCount}/${
					validImages.length
				} 張成功插入到 ${path.basename(excelPath)}`
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
			console.error(`❌ 圖片插入失敗：${path.basename(excelPath)} - ${error}`);
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
	 * 建立新的 Excel 檔案並同時插入資料和圖片
	 * @param {string} filePath - 輸出檔案路徑
	 * @param {any[][]} data - 要寫入的資料
	 * @param {ImageOptions[]} images - 圖片選項陣列
	 * @param {WriteOptions} options - 寫入選項
	 */
	async createExcelWithImages(
		filePath: string,
		data: any[][],
		images: ImageOptions[] = [],
		options: WriteOptions = {}
	): Promise<WriteResult> {
		try {
			// 先建立 Excel 檔案
			const writeResult = await this.writeExcelFile(filePath, data, options);

			if (!writeResult.success) {
				return writeResult;
			}

			// 如果有圖片要插入
			if (images.length > 0) {
				const imageResult = await this.insertImages(
					filePath,
					images,
					options.worksheetName
				);

				return {
					...writeResult,
					imagesInserted: imageResult.success ? imageResult.imagesInserted : 0,
					updatedAt: imageResult.updatedAt,
					error: !imageResult.success ? imageResult.error : writeResult.error,
				};
			}

			return writeResult;
		} catch (error) {
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
	 * 根據資料自動產生圖片插入配置
	 * @param {any[][]} data - Excel 資料
	 * @param {string} imageFolder - 圖片資料夾路徑
	 * @param {string} idColumn - ID 欄位名稱或索引
	 * @param {string} targetColumn - 目標欄位 (如: 'G')
	 * @param {object} imageSettings - 圖片設定
	 */
	generateImageConfigs(
		data: any[][],
		imageFolder: string,
		idColumn: number | string,
		targetColumn: string,
		imageSettings: Partial<ImageOptions> = {}
	): ImageOptions[] {
		const images: ImageOptions[] = [];

		// 跳過標題行，從第二行開始
		for (let i = 1; i < data.length; i++) {
			const row = data[i];
			let id: string;

			// 根據 idColumn 類型取得 ID
			if (typeof idColumn === "number") {
				id = row[idColumn]?.toString();
			} else {
				// 如果是字串，需要找到對應的欄位索引
				const headerRow = data[0];
				const columnIndex = headerRow.findIndex(
					(header: any) => header === idColumn
				);
				if (columnIndex === -1) {
					console.warn(`找不到欄位：${idColumn}`);
					continue;
				}
				id = row[columnIndex]?.toString();
			}

			if (!id) {
				console.warn(`第 ${i + 1} 行沒有 ID 值`);
				continue;
			}

			// 建構圖片路徑
			const imagePath = path.join(imageFolder, `${id}.png`);
			const cell = `${targetColumn}${i + 1}`;

			images.push({
				imagePath,
				cell,
				width: 80,
				height: 80,
				maintainAspectRatio: true,
				...imageSettings,
			});
		}

		return images;
	}

	/**
	 * 解析儲存格地址 (如: 'A1' -> {col: 1, row: 1})
	 * @private
	 */
	private parseCellAddress(cell: string): { col: number; row: number } {
		const match = cell.match(/^([A-Z]+)(\d+)$/);
		if (!match) {
			throw new Error(`無效的儲存格地址：${cell}`);
		}

		const colStr = match[1];
		const rowStr = match[2];

		// 將欄位字母轉換為數字 (A=1, B=2, ...)
		let col = 0;
		for (let i = 0; i < colStr.length; i++) {
			col = col * 26 + (colStr.charCodeAt(i) - "A".charCodeAt(0) + 1);
		}

		const row = parseInt(rowStr, 10);

		return { col, row };
	}

	/**
	 * 批次處理：插入多個工作表的圖片
	 * @param {string} excelPath - Excel 檔案路徑
	 * @param {Map<string, ImageOptions[]>} worksheetImages - 工作表名稱對應的圖片陣列
	 */
	async insertImagesMultipleWorksheets(
		excelPath: string,
		worksheetImages: Map<string, ImageOptions[]>
	): Promise<WriteResult[]> {
		const results: WriteResult[] = [];

		for (const [worksheetName, images] of worksheetImages) {
			const result = await this.insertImages(excelPath, images, worksheetName);
			results.push(result);
		}

		return results;
	}

	/**
	 * 安全地加載圖片到 ExcelJS
	 * @private
	 */
	private addImageToWorkbook(
		workbook: ExcelJS.Workbook,
		imagePath: string
	): number {
		try {
			const ext = path.extname(imagePath).toLowerCase().replace(".", "");

			// 方法一：使用 filename（推薦）
			return workbook.addImage({
				filename: imagePath,
				extension: ext as "jpeg" | "png" | "gif",
			});
		} catch (error) {
			throw new Error(`無法加載圖片：${imagePath} - ${error}`);
		}
	}

	async insertImagesSafely(
		excelPath: string,
		images: ImageOptions[],
		worksheetName?: string
	): Promise<WriteResult> {
		try {
			console.log(`🖼️ 正在安全插入 ${images.length} 張圖片到：${excelPath}`);

			if (!this.checkFileExists(excelPath)) {
				throw new Error(`Excel 檔案不存在：${excelPath}`);
			}

			// 檢查並過濾有效的圖片檔案
			const validImages: ImageOptions[] = [];
			for (const img of images) {
				if (this.checkFileExists(img.imagePath)) {
					// 檢查檔案是否可讀取
					try {
						const stats = fs.statSync(img.imagePath);
						if (stats.size > 0) {
							validImages.push(img);
						} else {
							console.warn(`⚠️ 圖片檔案為空，跳過：${img.imagePath}`);
						}
					} catch (statError) {
						console.warn(`⚠️ 無法讀取圖片檔案，跳過：${img.imagePath}`);
					}
				} else {
					console.warn(`⚠️ 圖片檔案不存在，跳過：${img.imagePath}`);
				}
			}

			if (validImages.length === 0) {
				console.warn("⚠️ 沒有有效的圖片檔案可插入");
				return {
					success: true,
					fileName: path.basename(excelPath),
					filePath: excelPath,
					imagesInserted: 0,
					updatedAt: new Date().toISOString(),
				};
			}

			const workbook = new ExcelJS.Workbook();
			await workbook.xlsx.readFile(excelPath);

			// 選擇工作表
			const worksheet = worksheetName
				? workbook.getWorksheet(worksheetName)
				: workbook.getWorksheet(1);

			if (!worksheet) {
				throw new Error(`找不到工作表：${worksheetName || "第一個工作表"}`);
			}

			// 批次插入圖片，每次插入後稍作延遲
			let insertedCount = 0;
			for (const [index, imageOption] of validImages.entries()) {
				try {
					const {
						imagePath,
						cell,
						width = 60,
						height = 60,
						maintainAspectRatio = true,
					} = imageOption;

					console.log(
						`📷 正在處理圖片 ${index + 1}/${
							validImages.length
						}: ${path.basename(imagePath)}`
					);

					// 使用更安全的圖片加載方法
					const imageId = await this.addImageToWorkbookSafely(
						workbook,
						imagePath
					);

					// 解析儲存格位置
					const cellInfo = this.parseCellAddress(cell);

					// 設定圖片位置和大小 - 使用更精確的定位
					const imageConfig: any = {
						tl: {
							col: cellInfo.col - 1,
							row: cellInfo.row - 1,
							colOff: 0.1, // 稍微偏移避免與儲存格邊緣重疊
							rowOff: 0.1,
						},
						ext: { width, height },
					};

					// 如果要保持長寬比
					if (maintainAspectRatio) {
						delete imageConfig.ext.height;
						imageConfig.ext.width = width;
					}

					worksheet.addImage(imageId, imageConfig);
					insertedCount++;

					console.log(`✅ 圖片插入成功：${path.basename(imagePath)} → ${cell}`);

					// 每處理 10 張圖片休息一下，避免系統負擔過重
					if ((index + 1) % 10 === 0) {
						await new Promise((resolve) => setTimeout(resolve, 100));
					}
				} catch (imgError) {
					console.error(
						`❌ 插入圖片失敗：${imageOption.imagePath} - ${imgError}`
					);
				}
			}

			// 儲存檔案前先備份
			const backupPath = excelPath.replace(".xlsx", "_backup.xlsx");
			try {
				fs.copyFileSync(excelPath, backupPath);
				console.log(`💾 已建立備份檔案：${backupPath}`);
			} catch (backupError) {
				console.warn(`⚠️ 無法建立備份檔案：${backupError}`);
			}

			// 分段儲存，降低錯誤風險
			try {
				await workbook.xlsx.writeFile(excelPath);
				console.log(
					`🎉 圖片插入完成：${insertedCount}/${validImages.length} 張成功插入`
				);
			} catch (saveError) {
				// 如果儲存失敗，嘗試另存新檔
				const altPath = excelPath.replace(".xlsx", "_with_images.xlsx");
				await workbook.xlsx.writeFile(altPath);
				console.log(`⚠️ 原檔案儲存失敗，已另存為：${altPath}`);
			}

			return {
				success: true,
				fileName: path.basename(excelPath),
				filePath: excelPath,
				imagesInserted: insertedCount,
				worksheetName: worksheet.name,
				updatedAt: new Date().toISOString(),
			};
		} catch (error) {
			console.error(
				`❌ 安全圖片插入失敗：${path.basename(excelPath)} - ${error}`
			);
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
	 * 更安全的圖片加載方法
	 * @private
	 */
	private async addImageToWorkbookSafely(
		workbook: ExcelJS.Workbook,
		imagePath: string
	): Promise<number> {
		const ext = path.extname(imagePath).toLowerCase().replace(".", "");

		// 確保副檔名有效
		let validExtension: "jpeg" | "png" | "gif";
		switch (ext) {
			case "jpg":
			case "jpeg":
				validExtension = "jpeg";
				break;
			case "png":
				validExtension = "png";
				break;
			case "gif":
				validExtension = "gif";
				break;
			default:
				throw new Error(`不支援的圖片格式: ${ext}`);
		}

		// 先嘗試使用 filename 方法
		try {
			// 使用絕對路徑
			const absolutePath = path.resolve(imagePath);

			return workbook.addImage({
				filename: absolutePath,
				extension: validExtension,
			});
		} catch (filenameError) {
			console.warn(`filename 方法失敗，嘗試 buffer 方法: ${filenameError}`);

			// 備用：使用 buffer 方法
			try {
				const imageBuffer = fs.readFileSync(imagePath);

				// 確保 buffer 不為空
				if (imageBuffer.length === 0) {
					throw new Error("圖片檔案為空");
				}

				// 安全的 buffer 轉換
				const safeBuffer = imageBuffer.buffer;

				return workbook.addImage({
					buffer: safeBuffer,
					extension: validExtension,
				});
			} catch (bufferError) {
				throw new Error(`圖片載入完全失敗 ${imagePath}: ${bufferError}`);
			}
		}
	}

	/**
	 * 檢查並修復 Excel 檔案
	 * @param excelPath Excel 檔案路徑
	 */
	async repairExcelFile(excelPath: string): Promise<boolean> {
		try {
			console.log("🔧 正在嘗試修復 Excel 檔案...");

			// 讀取檔案
			const workbook = new ExcelJS.Workbook();
			await workbook.xlsx.readFile(excelPath);

			// 重新儲存，這通常可以修復一些輕微的錯誤
			const repairedPath = excelPath.replace(".xlsx", "_repaired.xlsx");
			await workbook.xlsx.writeFile(repairedPath);

			console.log(`✅ Excel 檔案已修復並儲存為：${repairedPath}`);
			return true;
		} catch (error) {
			console.error(`❌ Excel 檔案修復失敗：${error}`);
			return false;
		}
	}
}
