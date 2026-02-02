import { IsOptional, IsString, MinLength } from 'class-validator';

export class CreatePlanDto {
  @IsString()
  @MinLength(1)
  name!: string;

  @IsString()
  @MinLength(1)
  year!: string;

  @IsOptional()
  @IsString()
  status?: string;
}
