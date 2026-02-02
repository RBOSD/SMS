import { IsOptional, IsString, MinLength } from 'class-validator';

export class CreatePlanScheduleDto {
  @IsString()
  @MinLength(1)
  railway!: string;

  @IsString()
  @MinLength(1)
  inspectionType!: string;

  @IsOptional()
  @IsString()
  planNumber?: string;

  @IsOptional()
  @IsString()
  startDate?: string; // ISO date (YYYY-MM-DD)

  @IsOptional()
  @IsString()
  endDate?: string; // ISO date (YYYY-MM-DD)

  @IsOptional()
  @IsString()
  business?: string;

  @IsOptional()
  @IsString()
  location?: string;

  @IsOptional()
  @IsString()
  inspector?: string;
}
