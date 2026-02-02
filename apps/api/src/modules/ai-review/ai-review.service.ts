import { Injectable, ServiceUnavailableException } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { GoogleGenerativeAI } from '@google/generative-ai';

@Injectable()
export class AiReviewService {
  constructor(private readonly config: ConfigService) {}

  async analyze(params: { content: string; handling: string }) {
    const apiKey = this.config.get<string>('GEMINI_API_KEY') || '';
    if (!apiKey) {
      throw new ServiceUnavailableException(
        'GEMINI_API_KEY 未設定，無法使用 AI 審查功能',
      );
    }

    const genAI = new GoogleGenerativeAI(apiKey);
    const model = genAI.getGenerativeModel({ model: 'gemini-1.5-flash' });

    const prompt = [
      '你是監理機關的審查人員，請依據「事項內容」與「機構辦理情形」產出審查意見（繁體中文）。',
      '輸出要求：',
      '1) 用條列列出重點（最多 6 點）',
      '2) 每點要具體、可執行、避免空泛',
      '',
      '【事項內容】',
      params.content,
      '',
      '【機構辦理情形】',
      params.handling,
    ].join('\n');

    const r = await model.generateContent(prompt);
    const text = r.response.text();
    return text;
  }
}
