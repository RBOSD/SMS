import { Module } from '@nestjs/common';
import { AuthModule } from '../auth/auth.module';
import { FeatureFlagsModule } from '../feature-flags/feature-flags.module';
import { AiReviewService } from './ai-review.service';
import { GeminiController } from './ai-review.controller';

@Module({
  imports: [AuthModule, FeatureFlagsModule],
  controllers: [GeminiController],
  providers: [AiReviewService],
})
export class AiReviewModule {}
