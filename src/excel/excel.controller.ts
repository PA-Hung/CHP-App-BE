import { Controller, Get, Post, Body, Patch, Param, Delete, UseInterceptors, UploadedFile, Res, Query } from '@nestjs/common';
import { ExcelService } from './excel.service';
import { UpdateExcelDto } from './dto/update-excel.dto';
import { FileInterceptor } from '@nestjs/platform-express';
import { ReqUser } from 'src/decorator/customize';
import { IUser } from 'src/users/users.interface';

@Controller('excel')
export class ExcelController {
  constructor(private readonly excelService: ExcelService) { }

  @Post('import')
  @UseInterceptors(FileInterceptor('fileExcel'))
  importExcel(@UploadedFile() file: Express.Multer.File, @ReqUser() userInfo: IUser) {
    return this.excelService.importExcel(file, userInfo);
  }

  @Get('export')
  exportExcel(@ReqUser() userInfo: IUser) {
    return this.excelService.exportExcel(userInfo);
  }

  @Get()
  findAll() {
    return this.excelService.findAll();
  }

  @Get(':id')
  findOne(@Param('id') id: string) {
    return this.excelService.findOne(+id);
  }

  @Patch(':id')
  update(@Param('id') id: string, @Body() updateExcelDto: UpdateExcelDto) {
    return this.excelService.update(+id, updateExcelDto);
  }

  @Delete(':id')
  remove(@Param('id') id: string) {
    return this.excelService.remove(+id);
  }
}
