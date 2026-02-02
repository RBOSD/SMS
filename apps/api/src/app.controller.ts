import { Controller, Get } from '@nestjs/common';

@Controller()
export class AppController {
  @Get()
  health() {
    return { ok: true, service: 'sms-v2-api', ts: new Date().toISOString() };
  }
}
