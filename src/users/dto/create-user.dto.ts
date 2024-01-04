import { IsNotEmpty } from 'class-validator';

export class CreateUserDto {
    @IsNotEmpty({ message: 'Name không được để trống !' })
    name: string;

    @IsNotEmpty({ message: 'Phone không được để trống !' })
    phone: string;

    @IsNotEmpty()
    password: string;

    @IsNotEmpty()
    role: string;
}

export class RegisterUserDto {
    @IsNotEmpty({
        message: 'Name không được để trống !',
    })
    name: string;

    @IsNotEmpty({
        message: 'Phone không được để trống !',
    })
    phone: string;

    @IsNotEmpty()
    password: string;
}
