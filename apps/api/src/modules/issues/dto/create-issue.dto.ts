import { IsInt, IsOptional, IsString, MinLength } from 'class-validator';

export class CreateIssueDto {
  @IsString()
  @MinLength(1)
  number!: string;

  @IsOptional()
  @IsString()
  year?: string;

  @IsOptional()
  @IsString()
  unit?: string;

  @IsOptional()
  @IsString()
  content?: string;

  @IsOptional()
  @IsString()
  status?: string;

  @IsOptional()
  @IsInt()
  planId?: number;
}
