import { PartialType } from '@nestjs/mapped-types';
import { CreateAccommodationDto } from './create-accommodation.dto';
import { IsNotEmpty } from 'class-validator';

export class UpdateAccommodationDto extends PartialType(CreateAccommodationDto) {
    @IsNotEmpty()
    _id: string
}
