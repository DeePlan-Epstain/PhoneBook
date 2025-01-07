import { createTheme } from '@material-ui/core/styles';
export const theme = createTheme({
    direction: 'rtl', // Both here and <body dir="rtl">
    palette: {
        primary: {
            main: '#0A3F67'
        },
        secondary: {
            main: '#FFA500'
          }

    }
});