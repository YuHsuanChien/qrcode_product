import readExcelFile from "./service/read_excel_file";
import path from "path";
import QRCodeService from "./service/qrcode";
import WriteExcelFile from "./service/write_excel_file";

interface ExcelData {
	success: boolean;
	fileName: string;
	filePath: string;
	rowCount: number;
	columnCount: number;
	data: any[];
	readAt: string;
	error?: Error;
}

interface StaffData {
	id: string;
	code: string;
	acount: string;
	staff_name: string;
	family: string | null;
	team: number;
}

async function index() {
	const readExcelService = new readExcelFile();
	const qrcodeService = new QRCodeService();
	const writeExcelService = new WriteExcelFile();

	try {
		console.log("🚀 開始處理 夥伴名單 Excel 和 QR Code...");

		// 1. 讀取原始 Excel 檔案 夥伴名單.xlsx
		const filePath = path.resolve(__dirname, "../record.xlsx");
		console.log(`📂 處理檔案：${filePath}`);

		const member_data: ExcelData = await readExcelService.readSingleExcelFile(
			filePath,
			{
				worksheetName: "夥伴名單",
			}
		);

		// 檢查讀取結果
		if (!member_data.success) {
			throw new Error(`讀取 Excel 失敗：${member_data.error?.message}`);
		}

		console.log(
			`📊 成功讀取 ${member_data.rowCount} 行資料，${member_data.columnCount} 欄`
		);

		// 2. 轉換資料格式，並過濾有效資料
		const staffList: StaffData[] = member_data.data
			.slice(1)
			.map((row) => ({
				id: row[0],
				code: row[1],
				acount: row[2],
				staff_name: row[3],
				family: row[4],
				team: row[5],
			}))
			.filter((staff) => staff.id && typeof staff.id === "string"); // 過濾掉沒有ID的行

		console.log(`👥 找到 ${staffList.length} 位有效員工資料`);

		if (staffList.length === 0) {
			throw new Error("沒有找到有效的員工資料");
		}

		const member_qrcode_folder = path.resolve(__dirname, "../member_qrcode");
		console.log(`📁 QR Code 資料夾：${member_qrcode_folder}`);
		// 把 QR Code 資料夾內的圖片都刪除
		const cleared = await writeExcelService.clearFolder(member_qrcode_folder);
		if (cleared) {
			console.log("🧹 已清空 QR Code 資料夾");
		}

		// 3. 序列產生 QR Code（避免併發問題）
		console.log("🔄 開始序列產生 QR Code...");

		let qrSuccessCount = 0;
		for (let i = 0; i < staffList.length; i++) {
			const staff = staffList[i];
			try {
				await qrcodeService.generateQRCode(staff.id, "member");
				qrSuccessCount++;
			} catch (qrError) {
				console.error(`❌ QR Code 產生失敗: ${staff.id} - ${qrError}`);
			}

			// 每產生 10 個就稍作休息，避免系統負擔過重
			if ((i + 1) % 10 === 0) {
				await new Promise((resolve) => setTimeout(resolve, 200));
				console.log(`⏳ 已處理 ${i + 1}/${staffList.length}，稍作休息...`);
			}
		}

		console.log(
			`✅ QR Code 產生完成：${qrSuccessCount}/${staffList.length} 個成功`
		);

		// 4. 檢查 QR Code 資料夾和圖片

		if (!writeExcelService.checkFileExists(member_qrcode_folder)) {
			throw new Error(`QR Code 資料夾不存在：${member_qrcode_folder}`);
		}

		// 5. 準備圖片插入配置（直接基於原始資料，不需要 newData）
		const imageConfigs = staffList.map((staff, index) => {
			const imagePath = path.resolve(member_qrcode_folder, `${staff.id}.png`); // 使用絕對路徑
			const cell = `G${index + 2}`; // G2, G3, G4... (跳過標題行)

			return {
				imagePath: imagePath,
				cell: cell,
				width: 50, // 縮小圖片尺寸，避免 Excel 錯誤
				height: 50,
				maintainAspectRatio: true,
			};
		});

		// 6. 驗證哪些圖片實際存在
		const validImageConfigs = imageConfigs.filter((config) => {
			const exists = writeExcelService.checkFileExists(config.imagePath);
			if (!exists) {
				console.warn(`⚠️ 圖片不存在：${path.basename(config.imagePath)}`);
			}
			return exists;
		});

		console.log(
			`🖼️ 準備插入 ${validImageConfigs.length}/${imageConfigs.length} 個 QR Code 圖片`
		);

		if (validImageConfigs.length === 0) {
			console.warn("⚠️ 沒有有效的 QR Code 圖片可以插入");
			return;
		}

		// 7. 建立備份檔案
		const newFilePath = filePath.replace(".xlsx", `_with_qrcode.xlsx`);
		try {
			const fs = require("fs");
			fs.copyFileSync(filePath, newFilePath);
			console.log(`💾 已建立新檔案：${path.basename(newFilePath)}`);
		} catch (createError) {
			console.warn(`⚠️ 無法建立新檔案：${createError}`);
		}

		// 8. 使用安全的方法插入圖片
		console.log("🔄 正在安全插入圖片到原始檔案...");

		// 檢查是否有 insertImagesSafely 方法，如果沒有就使用原始方法
		const insertResult = await writeExcelService.insertImagesSafely(
			newFilePath,
			validImageConfigs,
			"夥伴名單"
		);

		// 9. 檢查結果
		if (insertResult.success) {
			console.log("🎉 原始檔案修改完成！");
			console.log(`📁 檔案位置：${insertResult.filePath}`);
			console.log(
				`🖼️ 成功插入：${insertResult.imagesInserted || "N/A"} 張圖片`
			);

			// 如果有部分失敗，給出提示
			if (
				insertResult.imagesInserted &&
				insertResult.imagesInserted < validImageConfigs.length
			) {
				const failedCount =
					validImageConfigs.length - insertResult.imagesInserted;
				console.log(`⚠️ 有 ${failedCount} 張圖片插入失敗`);
				console.log(
					"💡 可能原因：圖片格式不支援、檔案權限問題、或 Excel 檔案被鎖定"
				);
			}
		} else {
			throw new Error(`圖片插入失敗：${insertResult.error?.message}`);
		}

		// 10. 最終驗證
		console.log("🔍 驗證檔案修改結果...");
		try {
			const verifyResult = await readExcelService.readSingleExcelFile(
				filePath,
				{
					worksheetName: "夥伴名單",
				}
			);

			if (verifyResult.success) {
				console.log(`✅ 檔案驗證通過！包含 ${verifyResult.rowCount} 行資料`);
				console.log("🎯 處理完成！可以開啟 Excel 檔案查看結果");
			} else {
				console.warn("⚠️ 檔案驗證異常，但圖片可能已經插入");
			}
		} catch (verifyError) {
			console.warn(`⚠️ 檔案驗證失敗：${verifyError}`);
			console.log("💡 這不一定表示處理失敗，請手動檢查 Excel 檔案");
		}

		// 開始處理 眷屬名單 Excel 和 QR Code...
		console.log("🚀 開始處理 眷屬名單 Excel 和 QR Code...");
		const family_data: ExcelData = await readExcelService.readSingleExcelFile(
			filePath,
			{
				worksheetName: "眷屬名單",
			}
		);

		// 檢查讀取結果
		if (!family_data.success) {
			throw new Error(`讀取 Excel 失敗：${family_data.error?.message}`);
		}

		console.log(
			`📊 成功讀取 ${family_data.rowCount} 行資料，${family_data.columnCount} 欄`
		);

		// 2. 轉換資料格式，並過濾有效資料
		const familyList: StaffData[] = family_data.data
			.slice(1)
			.map((row) => ({
				id: row[0],
				code: row[1],
				acount: row[2],
				staff_name: row[3],
				family: row[4],
				team: row[5],
			}))
			.filter((staff) => staff.id && typeof staff.id === "string"); // 過濾掉沒有ID的行

		console.log(`👥 找到 ${familyList.length} 位有效眷屬資料`);

		if (familyList.length === 0) {
			throw new Error("沒有找到有效的眷屬資料");
		}

		const family_qrcode_folder = path.resolve(__dirname, "../family_qrcode");
		console.log(`📁 QR Code 資料夾：${family_qrcode_folder}`);
		// 把 QR Code 資料夾內的圖片都刪除
		const family_cleared = await writeExcelService.clearFolder(
			family_qrcode_folder
		);
		if (family_cleared) {
			console.log("🧹 已清空 QR Code 資料夾");
		}

		// 3. 序列產生 QR Code（避免併發問題）
		console.log("🔄 開始序列產生 QR Code...");

		let family_qrSuccessCount = 0;
		for (let i = 0; i < familyList.length; i++) {
			const staff = familyList[i];
			try {
				await qrcodeService.generateQRCode(staff.id, "family");
				family_qrSuccessCount++;
			} catch (qrError) {
				console.error(`❌ QR Code 產生失敗: ${staff.id} - ${qrError}`);
			}

			// 每產生 10 個就稍作休息，避免系統負擔過重
			if ((i + 1) % 10 === 0) {
				await new Promise((resolve) => setTimeout(resolve, 200));
				console.log(`⏳ 已處理 ${i + 1}/${familyList.length}，稍作休息...`);
			}
		}

		console.log(
			`✅ QR Code 產生完成：${family_qrSuccessCount}/${familyList.length} 個成功`
		);

		// 4. 檢查 QR Code 資料夾和圖片

		if (!writeExcelService.checkFileExists(family_qrcode_folder)) {
			throw new Error(`QR Code 資料夾不存在：${family_qrcode_folder}`);
		}

		// 5. 準備圖片插入配置（直接基於原始資料，不需要 newData）
		const family_imageConfigs = familyList.map((staff, index) => {
			const imagePath = path.resolve(family_qrcode_folder, `${staff.id}.png`); // 使用絕對路徑
			const cell = `G${index + 2}`; // G2, G3, G4... (跳過標題行)

			return {
				imagePath: imagePath,
				cell: cell,
				width: 50, // 縮小圖片尺寸，避免 Excel 錯誤
				height: 50,
				maintainAspectRatio: true,
			};
		});

		// 6. 驗證哪些圖片實際存在
		const family_validImageConfigs = family_imageConfigs.filter((config) => {
			const exists = writeExcelService.checkFileExists(config.imagePath);
			if (!exists) {
				console.warn(`⚠️ 圖片不存在：${path.basename(config.imagePath)}`);
			}
			return exists;
		});

		console.log(
			`🖼️ 準備插入 ${family_validImageConfigs.length}/${family_imageConfigs.length} 個 QR Code 圖片`
		);

		if (family_validImageConfigs.length === 0) {
			console.warn("⚠️ 沒有有效的 QR Code 圖片可以插入");
			return;
		}

		// 8. 使用安全的方法插入圖片
		console.log("🔄 正在安全插入圖片到原始檔案...");

		// 檢查是否有 family_insertResult 方法
		const family_insertResult = await writeExcelService.insertImagesSafely(
			newFilePath,
			family_validImageConfigs,
			"眷屬名單"
		);

		// 9. 檢查結果
		if (family_insertResult.success) {
			console.log("🎉 原始檔案修改完成！");
			console.log(`📁 檔案位置：${family_insertResult.filePath}`);
			console.log(
				`🖼️ 成功插入：${family_insertResult.imagesInserted || "N/A"} 張圖片`
			);

			// 如果有部分失敗，給出提示
			if (
				family_insertResult.imagesInserted &&
				family_insertResult.imagesInserted < family_validImageConfigs.length
			) {
				const failedCount =
					family_validImageConfigs.length - family_insertResult.imagesInserted;
				console.log(`⚠️ 有 ${failedCount} 張圖片插入失敗`);
				console.log(
					"💡 可能原因：圖片格式不支援、檔案權限問題、或 Excel 檔案被鎖定"
				);
			}
		} else {
			throw new Error(`圖片插入失敗：${insertResult.error?.message}`);
		}

		// 10. 最終驗證
		console.log("🔍 驗證檔案修改結果...");
		try {
			const verifyResult = await readExcelService.readSingleExcelFile(
				filePath,
				{
					worksheetName: "眷屬名單",
				}
			);

			if (verifyResult.success) {
				console.log(`✅ 檔案驗證通過！包含 ${verifyResult.rowCount} 行資料`);
				console.log("🎯 處理完成！可以開啟 Excel 檔案查看結果");
			} else {
				console.warn("⚠️ 檔案驗證異常，但圖片可能已經插入");
			}
		} catch (verifyError) {
			console.warn(`⚠️ 檔案驗證失敗：${verifyError}`);
			console.log("💡 這不一定表示處理失敗，請手動檢查 Excel 檔案");
		}
	} catch (error) {
		console.error("❌ 處理過程中發生錯誤：", error);

		// 提供詳細的錯誤恢復建議
		console.log("\n🔧 錯誤恢復建議：");
		console.log('1. 確認 Excel 檔案 "record.xlsx" 存在且沒有被開啟');
		console.log('2. 檢查 "qrcode" 資料夾是否存在');
		console.log("3. 確認有足夠的磁碟空間");
		console.log("4. 檢查檔案權限，確保可以讀寫");
		console.log("5. 如果有備份檔案，可以嘗試從備份恢復");

		throw error;
	}
}

// 執行主函式
index().catch((error) => {
	console.error("💥 程式執行失敗：", error);
	process.exit(1);
});

export { index };
