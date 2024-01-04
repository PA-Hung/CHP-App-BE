import { Controller, Get, Post, Body, Patch, Param, Delete, Query } from '@nestjs/common';
import { UsersService } from './users.service';
import { CreateUserDto } from './dto/create-user.dto';
import { UpdateUserDto } from './dto/update-user.dto';
import { Public, ReqUser, ResponseMessage } from 'src/decorator/customize';
import { IUser } from './users.interface';

@Controller('users')
export class UsersController {
  constructor(private readonly usersService: UsersService) { }

  @ResponseMessage('Create a new user !')
  @Post()
  async create(@Body() createUserData: CreateUserDto, @ReqUser() userAuthInfo: IUser) {
    const newUser = await this.usersService.create(createUserData, userAuthInfo);
    return {
      _id: newUser?._id,
      createdAt: newUser?.createdAt
    }
  }

  @Public()
  @ResponseMessage('Get All user info!')
  @Get()
  findAll(
    @Query('current') currentPage: string,
    @Query('pageSize') limit: string,
    @Query() queryString: string
  ) {
    return this.usersService.findAll(+currentPage, +limit, queryString);
  }

  @Public()
  @ResponseMessage('Get a user info!')
  @Get(':id')
  findOne(@Param('id') id: string) {
    return this.usersService.findOne(id);
  }


  @ResponseMessage('Update a user !')
  @Patch()
  async update(@Body() updateUserData: UpdateUserDto, @ReqUser() user: IUser) {
    return await this.usersService.update(updateUserData, user);
  }

  @ResponseMessage('Detele a user !')
  @Delete(':id')
  async remove(@Param('id') id: string, @ReqUser() userAuthInfo: IUser) {
    return await this.usersService.remove(id, userAuthInfo);
  }
}
