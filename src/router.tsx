// Hash routing (PRD §17.3): deep links like /#/about survive a refresh on a static host.

import { createHashRouter } from 'react-router';
import { Layout } from './components/Layout.tsx';
import { About } from './routes/About.tsx';
import { Home } from './routes/Home.tsx';
import { LabRoute } from './routes/LabRoute.tsx';
import { NotFound } from './routes/NotFound.tsx';
import { RouteError } from './routes/RouteError.tsx';

export const router = createHashRouter([
  {
    element: <Layout />,
    errorElement: <RouteError />,
    children: [
      { index: true, element: <Home /> },
      {
        path: 'lab/:n',
        element: <LabRoute />,
      },
      { path: 'about', element: <About /> },
      { path: '*', element: <NotFound /> },
    ],
  },
]);
