import { Injectable, OnModuleInit, Logger } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { InjectModel } from '@nestjs/mongoose';
import { SoftDeleteModel } from 'soft-delete-plugin-mongoose';

import { User, UserDocument } from 'src/users/schemas/user.schema';
import { UsersService } from 'src/users/users.service';
import { ADMIN_ROLE, INIT_PERMISSIONS, USER_ROLE } from './sample';

@Injectable()
export class DatabasesService implements OnModuleInit {
    private readonly logger = new Logger(DatabasesService.name);
    constructor(
        @InjectModel(User.name)
        private userModel: SoftDeleteModel<UserDocument>,
        private configService: ConfigService,
        private userService: UsersService
    ) { }

    async onModuleInit() {
        const isInit = this.configService.get<string>("INIT_DATA");
        if (Boolean(isInit)) {

            const countUser = await this.userModel.count({});

            if (countUser === 0) {

                await this.userModel.insertMany([
                    {
                        name: "I'm admin",
                        phone: "0933634933",
                        password: this.userService.hassPassword(this.configService.get<string>("INIT_PASSWORD")),
                        role: "ADMIN"
                    },
                    {
                        name: "Phan Anh Khoa",
                        phone: "0377832014",
                        password: this.userService.hassPassword(this.configService.get<string>("INIT_PASSWORD")),
                        role: "USER"
                    },
                    {
                        name: "Thanh Châu",
                        phone: "0933048933",
                        password: this.userService.hassPassword(this.configService.get<string>("INIT_PASSWORD")),
                        role: "USER"
                    },
                ])
            }

            if (countUser > 0) {
                this.logger.verbose('>>> ALREADY INIT SAMPLE DATA...');
            }
        }

    }
}
