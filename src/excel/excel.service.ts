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
        //console.log('Processing Sheet:', sheetName);
        const sheet = workbook.Sheets[sheetName];
        //console.log('Sheet Data:', sheet);

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
        await this.AccommodationModel.create(mappedData);

      }

      return { success: true, message: 'Nhập file thành công !' };
    } catch (error) {
      return { success: false, error: 'Invalid Excel file' };
    }
  }

  async exportExcel() {
    try {
      // Retrieve data from MongoDB
      const accommodations = await this.AccommodationModel.find().exec();

      // Prepare worksheet data, handling dates appropriately
      const worksheetData = accommodations.map((accommodation) => ({
        'Họ và tên (*)': accommodation.name,
        'Ngày, tháng, năm sinh (*)': accommodation.birthday, // Convert Date to ISO 8601 string
        'Giới tính (*)': accommodation.gender,
        'CMND/ CCCD/ Số định danh (*)': accommodation.identification_number,
        'Số hộ chiếu (*)': accommodation.passport,
        'Giấy tờ khác (*)': accommodation.documents,
        'Số điện thoại': accommodation.phone,
        'Nghề nghiệp': accommodation.job,
        'Nơi làm việc': accommodation.workplace,
        'Dân tộc': accommodation.ethnicity,
        'Quốc tịch (*)': accommodation.nationality,
        'Địa chỉ – Quốc gia (*)': accommodation.country,
        'Địa chỉ – Tỉnh thành': accommodation.province,
        'Địa chỉ – Quận huyện': accommodation.district,
        'Địa chỉ – Phường xã': accommodation.ward,
        'Địa chỉ – Số nhà': accommodation.address,
        'Loại cư trú (*)': accommodation.residential_status,
        'Thời gian lưu trú (đến ngày) (*)': accommodation.arrival, // Convert Date to ISO 8601 string
        'Thời gian lưu trú (đi ngày)': accommodation.departure ? accommodation.departure : '', // Convert Date to ISO 8601 string or empty if departure is null
        'Lý do lưu trú': accommodation.reason,
        'Số phòng/Mã căn hộ': accommodation.apartment,
        // Add other fields as needed
      }));

      const ws = xlsx.utils.json_to_sheet(worksheetData);
      const wb = xlsx.utils.book_new();
      xlsx.utils.book_append_sheet(wb, ws, 'Sheet 1');
      const buffer = xlsx.write(wb, { bookType: 'xlsx', type: 'buffer' });
      return buffer; //Assuming you have a filename set in the response object

    } catch (error) {
      return { success: false, error: 'Error exporting Excel file' };
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
