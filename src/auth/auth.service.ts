import { BadRequestException, Injectable } from '@nestjs/common';
import { UsersService } from 'src/users/users.service';
import { JwtService } from '@nestjs/jwt';
import { IUser } from 'src/users/users.interface';
import { RegisterUserDto } from 'src/users/dto/create-user.dto';
import { ConfigService } from '@nestjs/config';
import { Response } from 'express';
import ms from 'ms';


@Injectable()
export class AuthService {
    constructor(
        private usersService: UsersService,
        private jwtService: JwtService,
        private configService: ConfigService,
    ) { }

    async validateUser(phone: string, pass: string): Promise<any> {
        const user = await this.usersService.findOneByUsername(phone);
        if (user) {
            const isValidPassword = this.usersService.checkUserPassword(pass, user.password)
            if (isValidPassword === true) {
                return user;
            }
        }
        return null;
    }

    async loginService(user: IUser, response: Response) {
        const { _id, name, phone, role } = user;
        const payload = {
            sub: "token login",
            iss: "from server",
            _id,
            name,
            phone,
            role
        };

        const refresh_token = this.createRefreshToken(payload)
        // update user with refresh token
        await this.usersService.updateUserToken(refresh_token, _id)
        // set refresh token as cookies
        response.cookie('refresh_token', refresh_token, {
            httpOnly: true,
            maxAge: ms(this.configService.get<string>('JWT_REFRESH_EXPIRE')) // ms để chuyển đổi sang mili giây
        })

        return {
            access_token: this.jwtService.sign(payload),
            user: {
                _id,
                name,
                phone,
                role
            }
        };
    }

    async registerService(registerUserInfo: RegisterUserDto) {
        const newUser = await this.usersService.register(registerUserInfo)
        return {
            _id: newUser?._id,
            createdAt: newUser?.createdAt,
        };
    }

    createRefreshToken = (payload: any) => {
        const refreshToken = this.jwtService.sign(payload, {
            secret: this.configService.get<string>('JWT_REFRESH_TOKEN_SECRET'),
            expiresIn: this.configService.get<string>('JWT_REFRESH_EXPIRE'),
        })
        return refreshToken
    }

    refreshTokenService = async (refreshToken: string, response: Response) => {
        try {
            this.jwtService.verify(refreshToken, {
                secret: this.configService.get<string>('JWT_REFRESH_TOKEN_SECRET')
            })
            const user = await this.usersService.findUserByToken(refreshToken)
            if (user) {
                const { _id, name, phone, role } = user;
                const payload = {
                    sub: "refresh token",
                    iss: "from server",
                    _id,
                    name,
                    phone,
                    role
                };

                const refresh_token = this.createRefreshToken(payload)
                // update user with refresh token
                await this.usersService.updateUserToken(refresh_token, _id.toString())
                // set refresh token as cookies
                response.clearCookie('refresh_token')
                response.cookie('refresh_token', refresh_token, {
                    httpOnly: true,
                    maxAge: ms(this.configService.get<string>('JWT_REFRESH_EXPIRE')) // ms để chuyển đổi sang mili giây
                })

                return {
                    access_token: this.jwtService.sign(payload),
                    user: {
                        _id,
                        name,
                        phone,
                        role
                    }
                };
            } else {
                throw new BadRequestException(`Refresh token không hợp lệ !`);
            }
        } catch (error) {
            throw new BadRequestException(`Refresh token không hợp lệ !`);
        }
    }

    logoutService = async (userAuthInfo: IUser, response: Response) => {
        await this.usersService.updateUserToken(null, userAuthInfo._id)
        response.clearCookie('refresh_token')
        return ('ok')
    }
}
