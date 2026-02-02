import { Body, Controller, Post, UseGuards } from '@nestjs/common';
import { AdminOrManagerGuard } from '../../common/admin-or-manager.guard';
import { RequireFeature } from '../../common/feature-flag.decorator';
import { FeatureFlagGuard } from '../../common/feature-flag.guard';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import { RunAiDto } from './dto/run-ai.dto';
import { AiReviewService } from './ai-review.service';

@Controller('gemini')
@RequireFeature('module_ai_review')
@UseGuards(JwtAuthGuard, FeatureFlagGuard, AdminOrManagerGuard)
export class GeminiController {
  constructor(private readonly ai: AiReviewService) {}

  @Post()
  async run(@Body() body: RunAiDto) {
    const content = String(body.content || '').trim();
    const rounds = Array.isArray(body.rounds) ? body.rounds : [];
    const handling = String(rounds?.[0]?.handling || '').trim();

    const result = await this.ai.analyze({ content, handling });
    return {
      result,
    };
  }
}
