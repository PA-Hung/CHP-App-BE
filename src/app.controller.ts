import { Controller, Get, Post, Render, Request, UseGuards } from '@nestjs/common';
import { AppService } from './app.service';
import { ConfigService } from '@nestjs/config';
import { AuthService } from './auth/auth.service';

@Controller()
export class AppController {
  constructor(
    private readonly appService: AppService,
    private configService: ConfigService,
  ) { }

  // hiển thị trang chủ để test HomePage cho backend
  @Get()
  @Render("home.ejs")
  handleHomePage() {
    console.log(">>>>>>>> check port :", this.configService.get<string>("PORT"));
    const message = this.appService.getHello()
    return {
      message: message
    }
  }

}
