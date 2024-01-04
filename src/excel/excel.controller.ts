import { Controller, Get, Post, Body, Patch, Param, Delete, UseInterceptors, UploadedFile, Res } from '@nestjs/common';
import { ExcelService } from './excel.service';
import { UpdateExcelDto } from './dto/update-excel.dto';
import { FileInterceptor } from '@nestjs/platform-express';

@Controller('excel')
export class ExcelController {
  constructor(private readonly excelService: ExcelService) { }

  @Post('import')
  @UseInterceptors(FileInterceptor('fileExcel'))
  importExcel(@UploadedFile() file: Express.Multer.File) {
    return this.excelService.importExcel(file);
  }

  @Get('export')
  exportExcel() {
    return this.excelService.exportExcel();
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
