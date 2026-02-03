import { Type } from 'class-transformer';
import { IsBoolean, IsOptional, ValidateNested } from 'class-validator';

export class FeatureFlagsPatchDto {
  @IsOptional()
  @IsBoolean()
  module_issues?: boolean;

  @IsOptional()
  @IsBoolean()
  module_plans?: boolean;

  @IsOptional()
  @IsBoolean()
  module_ai_review?: boolean;
}

export class UpdateFeatureFlagsDto {
  @ValidateNested()
  @Type(() => FeatureFlagsPatchDto)
  flags!: FeatureFlagsPatchDto;
}
