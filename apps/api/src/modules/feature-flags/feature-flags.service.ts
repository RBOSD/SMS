import { Injectable, OnModuleInit } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import { AuthService } from '../auth/auth.service';
import { FEATURE_KEYS, type FeatureKey } from './feature-keys';

type FlagMap = Record<string, boolean>;

@Injectable()
export class FeatureFlagsService implements OnModuleInit {
  constructor(
    private readonly prisma: PrismaService,
    private readonly auth: AuthService,
  ) {}

  private cache: { at: number; flags: FlagMap } | null = null;
  private readonly ttlMs = 5_000;

  async onModuleInit() {
    // Seed best-effort (don't crash app startup if DB not ready yet)
    try {
      await this.ensureSeed();
    } catch {
      // ignore
    }
  }

  private invalidateCache() {
    this.cache = null;
  }

  private async loadFlagsMap(): Promise<FlagMap> {
    const now = Date.now();
    if (this.cache && now - this.cache.at < this.ttlMs) return this.cache.flags;

    const rows = await this.prisma.featureFlag.findMany();
    const map: FlagMap = {};
    for (const k of FEATURE_KEYS) map[k] = true;
    for (const r of rows) map[r.key] = r.enabled;

    this.cache = { at: now, flags: map };
    return map;
  }

  async ensureSeed() {
    await this.prisma.$transaction(
      FEATURE_KEYS.map((key) =>
        this.prisma.featureFlag.upsert({
          where: { key },
          update: {},
          create: { key, enabled: true },
        }),
      ),
    );
    this.invalidateCache();
  }

  async getAll(): Promise<Record<FeatureKey, boolean>> {
    const m = await this.loadFlagsMap();
    return {
      module_issues: !!m.module_issues,
      module_plans: !!m.module_plans,
      module_ai_review: !!m.module_ai_review,
    };
  }

  async update(patch: Partial<Record<FeatureKey, boolean>>) {
    const updates = FEATURE_KEYS.filter((k) => patch[k] != null).map((key) =>
      this.prisma.featureFlag.upsert({
        where: { key },
        update: { enabled: patch[key] === true },
        create: { key, enabled: patch[key] === true },
      }),
    );
    if (updates.length > 0) await this.prisma.$transaction(updates);
    this.invalidateCache();
    return await this.getAll();
  }

  async getEffectiveForUser(
    userId: number,
  ): Promise<Record<FeatureKey, boolean>> {
    const isAdmin = await this.auth.isAdminUser(userId);
    if (isAdmin) {
      return {
        module_issues: true,
        module_plans: true,
        module_ai_review: true,
      };
    }
    return await this.getAll();
  }

  async isEnabledForUser(userId: number, key: string): Promise<boolean> {
    const effective = await this.getEffectiveForUser(userId);
    return (effective as any)[key] === true;
  }
}
