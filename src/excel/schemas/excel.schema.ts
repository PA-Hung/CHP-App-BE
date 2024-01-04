import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import mongoose, { HydratedDocument } from 'mongoose';

export type ExcelDocument = HydratedDocument<Excel>;

@Schema({ timestamps: true })
export class Excel {
}

export const ExcelSchema = SchemaFactory.createForClass(Excel);