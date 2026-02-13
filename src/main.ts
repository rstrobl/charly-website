import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';

async function bootstrap() {
  const app = await NestFactory.create(AppModule);
  const port = process.env.PORT || 8080;
  await app.listen(port, '127.0.0.1');
  console.log(`🚀 Server running on http://localhost:${port}`);
}
bootstrap();
