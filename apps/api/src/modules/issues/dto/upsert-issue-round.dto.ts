import { IsInt, IsOptional, IsString, Min } from 'class-validator';

export class UpsertIssueRoundDto {
  @IsInt()
  @Min(1)
  round!: number;

  @IsOptional()
  @IsString()
  handling?: string;

  @IsOptional()
  @IsString()
  review?: string;

  @IsOptional()
  @IsString()
  replyDate?: string;

  @IsOptional()
  @IsString()
  responseDate?: string;
}
