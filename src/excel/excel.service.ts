import { Injectable } from '@nestjs/common';
import { CreateExcelDto } from './dto/create-excel.dto';
import { UpdateExcelDto } from './dto/update-excel.dto';
import * as xlsx from 'xlsx';
import { InjectModel } from '@nestjs/mongoose';
import { SoftDeleteModel } from 'soft-delete-plugin-mongoose';
import { Accommodation, AccommodationDocument } from 'src/accommodation/schemas/accommodation.schema';

@Injectable()
export class ExcelService {

  constructor(
    @InjectModel(Accommodation.name)
    private AccommodationModel: SoftDeleteModel<AccommodationDocument>,
  ) { }

  async importExcel(file: Express.Multer.File) {
    try {
      function convertToISO8601Date(inputDate: string): Date | null {
        const [day, month, year] = inputDate.split('/');
        if (!(day && month)) {
          return null;
        }
        const formattedDate = year ? new Date(`${year}-${month}-${day}T00:00:00.000Z`) : new Date(`${new Date().getFullYear()}-${month}-${day}T00:00:00.000Z`);
        return formattedDate;
      }

      const workbook = xlsx.read(file.buffer, { type: 'buffer' });
      const sheetNames = workbook.SheetNames;

      for (const sheetName of sheetNames) {
        const sheet = workbook.Sheets[sheetName];
        const rawData = xlsx.utils.sheet_to_json(sheet, { header: 2 });
        rawData.shift();

        const mappedData = rawData.map((item) => ({
          name: item['Họ và tên (*)'] || '',
          birthday: convertToISO8601Date(item['Ngày, tháng, năm sinh (*)']) || null,
          gender: item['Giới tính (*)'] || '',
          identification_number: item['CMND/ CCCD/ Số định danh (*)'] || '',
          passport: item['Số hộ chiếu (*)'] || '',
          documents: item['Giấy tờ khác (*)'] || '',
          phone: item['Số điện thoại'] || '',
          job: item['Nghề nghiệp'] || '',
          workplace: item['Nơi làm việc'] || '',
          ethnicity: item['Dân tộc'] || '',
          nationality: item['Quốc tịch (*)'] || '',
          country: item['Địa chỉ – Quốc gia (*)'] || '',
          province: item['Địa chỉ – Tỉnh thành'] || '',
          district: item['Địa chỉ – Quận huyện'] || '',
          ward: item['Địa chỉ – Phường xã'] || '',
          address: item['Địa chỉ – Số nhà'] || '',
          residential_status: item['Loại cư trú (*)'] || '',
          arrival: convertToISO8601Date(item['Thời gian lưu trú (đến ngày) (*)']) || null,
          departure: convertToISO8601Date(item['Thời gian lưu trú (đi ngày)']) || null,
          reason: item['Lý do lưu trú'] || '',
          apartment: item['Số phòng/Mã căn hộ'] || '',
        }));

        // Process the 'mappedData' array as needed
        const resData = await this.AccommodationModel.create(mappedData);
      }

      return { success: true, message: 'Import completed' };
    } catch (error) {
      return { success: false, error: 'Invalid Excel file' };
    }
  }

  create(createExcelDto: CreateExcelDto) {
    return 'This action adds a new excel';
  }

  findAll() {
    return `This action returns all excel`;
  }

  findOne(id: number) {
    return `This action returns a #${id} excel`;
  }

  update(id: number, updateExcelDto: UpdateExcelDto) {
    return `This action updates a #${id} excel`;
  }

  remove(id: number) {
    return `This action removes a #${id} excel`;
  }
}
