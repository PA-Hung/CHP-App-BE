import { Module } from '@nestjs/common';
import { ExcelService } from './excel.service';
import { ExcelController } from './excel.controller';
import { MongooseModule } from '@nestjs/mongoose';
import { Excel, ExcelSchema } from './schemas/excel.schema';
import { Accommodation, AccommodationSchema } from 'src/accommodation/schemas/accommodation.schema';

@Module({
  imports: [MongooseModule.forFeature([
    { name: Accommodation.name, schema: AccommodationSchema },
  ])],
  controllers: [ExcelController],
  providers: [ExcelService]
})
export class ExcelModule { }
