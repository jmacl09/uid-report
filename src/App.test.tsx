import React from 'react';
import { render, screen } from '@testing-library/react';
import UIDLookup from './pages/UIDLookup';

// Mock react-router-dom to avoid ESM import issues and provide minimal hooks used by UIDLookup
jest.mock('react-router-dom', () => ({
  useLocation: () => ({ search: '' }),
  useNavigate: () => jest.fn(),
}), { virtual: true });

test('renders UID Assistant landing view', () => {
  render(<UIDLookup />);
  // Expect the landing banner title to be present (exact match)
  expect(screen.getByText(/^UID Assistant$/i)).toBeInTheDocument();
});
