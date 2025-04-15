import { Module } from '@nestjs/common';
import { TypeOrmModule } from '@nestjs/typeorm';
import { ConfigModule, ConfigService } from '@nestjs/config';

@Module({
  imports: [
    TypeOrmModule.forRootAsync({
      imports: [ConfigModule],
      inject: [ConfigService],
      useFactory: (cfg: ConfigService) => {
        return {
          type: 'postgres',
          host: cfg.get('DB_HOST'),
          port: cfg.get('DB_PORT'),
          username: cfg.get('DB_ADMIN_LOGIN'),
          password: cfg.get('DB_ADMIN_PASSWORD'),
          database: cfg.get('DB_DATABASE_NAME'),
          synchronize: false,
          entities: [__dirname + '/../**/*.entity.{js,ts}'],
          migrations: ['dist/migration/*{.ts,.js}'],
        };
      },
    }),
  ],
})
export class DatabaseModule {}
