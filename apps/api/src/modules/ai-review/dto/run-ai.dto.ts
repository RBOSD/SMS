import { Type } from 'class-transformer';
import {
  IsArray,
  IsOptional,
  IsString,
  MinLength,
  ValidateNested,
} from 'class-validator';

export class RunAiDto {
  @IsString()
  @MinLength(1)
  content!: string;

  @IsArray()
  @ValidateNested({ each: true })
  @Type(() => AiRoundDto)
  rounds!: AiRoundDto[];
}

export class AiRoundDto {
  @IsOptional()
  @IsString()
  handling?: string;

  @IsOptional()
  @IsString()
  review?: string;
}
