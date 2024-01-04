import { Module } from '@nestjs/common';
import { AccommodationService } from './accommodation.service';
import { AccommodationController } from './accommodation.controller';
import { MongooseModule } from '@nestjs/mongoose';
import { Accommodation, AccommodationSchema } from './schemas/accommodation.schema';

@Module({
  imports: [MongooseModule.forFeature([
    { name: Accommodation.name, schema: AccommodationSchema },
  ])],
  controllers: [AccommodationController],
  providers: [AccommodationService]
})
export class AccommodationModule { }
