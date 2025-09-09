import QRCode from "qrcode";
import path from "path";
import fs from "fs";

export default class QRCodeService {
	async generateQRCode(id: string): Promise<string> {
		try {
			// 取得專案根目錄下的 qrcode 資料夾
			const dir = path.resolve(__dirname, "../../qrcode");
			// 若資料夾不存在則建立
			if (!fs.existsSync(dir)) {
				fs.mkdirSync(dir, { recursive: true });
			}
			const filePath = path.join(dir, `${id}.png`);
			const qrCode = await QRCode.toDataURL(id);
			await QRCode.toFile(filePath, id);
			return qrCode;
		} catch (error) {
			console.error("Error generating QR code:", error);
			throw new Error("Failed to generate QR code");
		}
	}
}
