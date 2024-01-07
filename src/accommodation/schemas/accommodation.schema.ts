import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import mongoose, { HydratedDocument } from 'mongoose';

export type AccommodationDocument = HydratedDocument<Accommodation>;

@Schema({ timestamps: true })
export class Accommodation {
    @Prop({ type: mongoose.Schema.Types.ObjectId, ref: 'User' })
    userId: mongoose.Schema.Types.ObjectId;
    @Prop()
    name: string
    @Prop()
    birthday: string
    @Prop()
    gender: string
    @Prop()
    identification_number: string
    @Prop()
    passport: string
    @Prop()
    documents: string
    @Prop()
    phone: string
    @Prop()
    job: string
    @Prop()
    workplace: string
    @Prop()
    ethnicity: string
    @Prop()
    nationality: string
    @Prop()
    country: string
    @Prop()
    province: string
    @Prop()
    district: string
    @Prop()
    ward: string
    @Prop()
    address: string
    @Prop()
    residential_status: string
    @Prop()
    arrival: string
    @Prop()
    departure: string
    @Prop()
    reason: string
    @Prop()
    apartment: string

    @Prop({ type: Object })
    createdBy: {
        _id: mongoose.Schema.Types.ObjectId,
        phone: string
    };

    @Prop({ type: Object })
    updatedBy: {
        _id: mongoose.Schema.Types.ObjectId,
        phone: string
    };

    @Prop({ type: Object })
    deletedBy: {
        _id: mongoose.Schema.Types.ObjectId,
        phone: string
    };

    @Prop()
    createdAt: Date;

    @Prop()
    updatedAt: Date;

    @Prop()
    isDeleted: boolean;

    @Prop()
    deletedAt: Date;
}

export const AccommodationSchema = SchemaFactory.createForClass(Accommodation);