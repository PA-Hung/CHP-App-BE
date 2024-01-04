import { BadRequestException, Injectable } from '@nestjs/common';
import { CreateUserDto, RegisterUserDto } from './dto/create-user.dto';
import { UpdateUserDto } from './dto/update-user.dto';
import { InjectModel } from '@nestjs/mongoose';
import mongoose from 'mongoose';
import { User, UserDocument } from './schemas/user.schema';
import { genSaltSync, hashSync, compareSync } from 'bcryptjs';
import { SoftDeleteModel } from 'soft-delete-plugin-mongoose';
import { IUser } from './users.interface';
import aqp from 'api-query-params';


@Injectable()
export class UsersService {
  constructor(
    @InjectModel(User.name)
    private userModel: SoftDeleteModel<UserDocument>,
  ) { }

  hassPassword = (inputPassword: string) => {
    const salt = genSaltSync(10);
    const hash = hashSync(inputPassword, salt);
    return hash
  }

  async create(createUserData: CreateUserDto, userAuthInfo: IUser) {
    const checkPhone = await this.userModel.findOne({ phone: createUserData.phone })
    const checkName = await this.userModel.findOne({ name: createUserData.name })
    if (checkPhone) {
      throw new BadRequestException(`Số điện thoại : ${createUserData.phone} đã tồn tại !`);
    }
    if (checkName) {
      throw new BadRequestException(`Tên : ${createUserData.name} đã tồn tại !`);
    }
    const hassPassword = this.hassPassword(createUserData.password)
    const userData = await this.userModel.create({
      phone: createUserData.phone,
      password: hassPassword,
      name: createUserData.name,
      role: createUserData.role,
      createdBy: {
        _id: userAuthInfo._id,
        phone: userAuthInfo.phone
      }
    })
    return userData
  }

  async register(registerUserInfo: RegisterUserDto) {
    const checkPhone = await this.userModel.findOne({ phone: registerUserInfo.phone })
    if (checkPhone) {
      throw new BadRequestException(`Số điện thoại : ${registerUserInfo.phone} đã tồn tại !`);
    }
    const hassPassword = this.hassPassword(registerUserInfo.password)
    const userData = await this.userModel.create({
      name: registerUserInfo.name,
      phone: registerUserInfo.phone,
      password: hassPassword,
      role: "USER"
    })
    return userData;
  }

  async findAll(currentPage: number, limit: number, queryString: string) {
    const { filter, projection, sort, population } = aqp(queryString);
    delete filter.current
    delete filter.pageSize
    let offset = (+currentPage - 1) * (+limit);
    let defaultLimit = +limit ? +limit : 10;
    const totalItems = (await this.userModel.find(filter)).length;
    const totalPages = Math.ceil(totalItems / defaultLimit);
    const result = await this.userModel.find(filter)
      .skip(offset)
      .limit(defaultLimit)
      .sort(sort as any)
      .populate(population)
      .select('-password')
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


  findOne(id: string) {
    if (!mongoose.Types.ObjectId.isValid(id)) {
      return "User not found !"
    }
    return this.userModel.findOne({ _id: id }).select('-password')
      .populate({ path: "role", select: { _id: 1, name: 1 } })
  }

  findOneByUsername(phone: string) {
    return this.userModel.findOne({ phone })
  }

  checkUserPassword(password: string, hashPass: string) {
    return compareSync(password, hashPass)
  }

  async update(updateUserData: UpdateUserDto, user: IUser) {
    const checkName = await this.userModel.findOne({ name: updateUserData.name });

    if (checkName && checkName._id.toString() !== updateUserData._id) {
      throw new BadRequestException(`Tên : ${updateUserData.name} đã tồn tại!`);
    }

    const checkPhone = await this.userModel.findOne({ phone: updateUserData.phone });

    if (checkPhone && checkPhone._id.toString() !== updateUserData._id) {
      throw new BadRequestException(`Số điện thoại : ${updateUserData.phone} đã tồn tại!`);
    }
    const updated = await this.userModel.updateOne(
      { _id: updateUserData._id },
      {
        ...updateUserData,
        updatedBy: {
          _id: user._id,
          phone: user.phone
        }
      }
    );
    return updated
  }

  async remove(id: string, userAuthInfo: IUser) {
    if (!mongoose.Types.ObjectId.isValid(id)) {
      return "User not found !"
    }
    //check xem có phải admin không ? nếu là admin thì không cho xoá
    const checkAdmin = await this.userModel.findOne({ _id: id })
    if (checkAdmin && checkAdmin.phone === "0933634933") {
      throw new BadRequestException(
        `Bạn không được xoá tài khoản admin !`
      );
    }

    await this.userModel.updateOne({ _id: id }, {
      deletedBy: {
        _id: userAuthInfo._id,
        phone: userAuthInfo.phone
      }
    });
    return this.userModel.softDelete({ _id: id })
  }

  updateUserToken = async (refreshToken: string, _id: string) => {
    return await this.userModel.updateOne({ _id }, { refreshToken });
  }

  findUserByToken = async (refreshToken: string) => {
    return await this.userModel.findOne({ refreshToken })
      .populate({ path: "role", select: { name: 1 } })
  }
}
