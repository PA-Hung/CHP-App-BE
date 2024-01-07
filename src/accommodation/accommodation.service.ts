import { BadRequestException, Injectable } from '@nestjs/common';
import { CreateAccommodationDto } from './dto/create-accommodation.dto';
import { UpdateAccommodationDto } from './dto/update-accommodation.dto';
import { InjectModel } from '@nestjs/mongoose';
import { Accommodation, AccommodationDocument } from './schemas/accommodation.schema';
import { SoftDeleteModel } from 'soft-delete-plugin-mongoose';
import { IUser } from 'src/users/users.interface';
import aqp from 'api-query-params';
import mongoose, { Types } from 'mongoose';

@Injectable()
export class AccommodationService {

  constructor(
    @InjectModel(Accommodation.name)
    private accommodationModel: SoftDeleteModel<AccommodationDocument>,
  ) { }

  async create(createAccommodationDto: CreateAccommodationDto, userAuthInfo: IUser) {
    const resData = await this.accommodationModel.create({
      userId: userAuthInfo._id,
      ...createAccommodationDto,
      createdBy: {
        _id: userAuthInfo._id,
        phone: userAuthInfo.phone
      }
    })
    return resData
  }

  async findAll(currentPage: number, limit: number, queryString: string) {
    const { filter, projection, sort, population } = aqp(queryString);

    // Kiểm tra xem có trường userId trong filter không
    if (filter.userId) {
      //console.log('filter.userId', filter.userId);

      //Chuyển nó thành String và Xoá bỏ / ở đầu và /i ở cuối (nếu có)
      filter.userId = String(filter.userId).replace(/^\/|\/i$/g, '');

      // Chuyển đổi thành ObjectId nếu giá trị là một chuỗi ObjectId hợp lệ
      if (Types.ObjectId.isValid(filter.userId)) {
        filter.userId = new Types.ObjectId(filter.userId);
      }
    }

    delete filter.current
    delete filter.pageSize
    let offset = (+currentPage - 1) * (+limit);
    let defaultLimit = +limit ? +limit : 10;
    const totalItems = (await this.accommodationModel.find(filter)).length;
    const totalPages = Math.ceil(totalItems / defaultLimit);
    const result = await this.accommodationModel.find(filter)
      .skip(offset)
      .limit(defaultLimit)
      .sort(sort as any)
      .populate(population)
      .select('')
      .exec();
    return {
      meta: {
        current: currentPage, //trang hiện tại
        pageSize: limit, //số lượng bản ghi đã lấy
        pages: totalPages, //tổng số trang với điều kiện query
        total: totalItems // tổng số phần tử (số bản ghi)
      },
      result //kết quả query
    }
  }

  findOne(id: number) {
    return `This action returns a #${id} accommodation`;
  }

  async update(updateAccommodationDto: UpdateAccommodationDto, userInfo: IUser) {
    const updated = await this.accommodationModel.updateOne(
      { _id: updateAccommodationDto._id },
      {
        ...updateAccommodationDto,
        updatedBy: {
          _id: userInfo._id,
          phone: userInfo.phone
        }
      }
    );
    return updated
  }

  async remove(id: string, userAuthInfo: IUser) {
    if (!mongoose.Types.ObjectId.isValid(id)) {
      return "Lưu trú không tồn tại !"
    }
    await this.accommodationModel.updateOne({ _id: id }, {
      deletedBy: {
        _id: userAuthInfo._id,
        phone: userAuthInfo.phone
      }
    });
    return this.accommodationModel.deleteOne({ _id: id })
  }
}
