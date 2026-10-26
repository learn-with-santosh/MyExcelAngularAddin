import { bootstrapApplication } from '@angular/platform-browser';
import { appConfig } from './app/app.config';
import { AppComponent } from './app/app.component';

const bootstrap = () => {
  bootstrapApplication(AppComponent, appConfig)
    .catch((err) => console.error(err));
};

// Check if Office.js is available and initialize before bootstrapping Angular
if (typeof Office !== 'undefined') {
  Office.onReady().then(() => {
    bootstrap();
  }).catch((err) => {
    console.error('Office.onReady error:', err);
    bootstrap();
  });
} else {
  // Direct browser testing without Office container
  bootstrap();
}
