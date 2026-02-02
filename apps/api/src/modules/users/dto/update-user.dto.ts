import {
  ArrayUnique,
  IsArray,
  IsIn,
  IsInt,
  IsOptional,
  IsString,
  MinLength,
} from 'class-validator';

export class UpdateUserDto {
  @IsOptional()
  @IsString()
  @MinLength(1)
  name?: string;

  @IsOptional()
  @IsString()
  @MinLength(8)
  password?: string;

  @IsOptional()
  @IsIn(['MANAGER', 'VIEWER'])
  role?: 'MANAGER' | 'VIEWER';

  @IsOptional()
  @IsArray()
  @ArrayUnique()
  @IsInt({ each: true })
  groupIds?: number[];
}
