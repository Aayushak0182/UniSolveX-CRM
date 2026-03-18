import React from 'react';
import './App.css';
import { Box, Drawer, List, ListItem, ListItemAvatar, Avatar, ListItemText, Typography, TextField, Divider } from '@mui/material';

const mockContacts = [
  { name: 'Vishal Jais', status: 'Expert', id: '307048' },
  { name: 'Rushma', status: 'Expert', id: 'EDU5175572' },
  { name: 'Priyanka', status: 'Expert', id: '509200' },
  { name: 'Ray', status: 'Expert', id: 'EDU5175564' },
  { name: 'Jiyani Nency', status: 'Expert', id: '507512' },
  { name: 'Ankush Tr', status: 'Expert', id: '509196' },
  { name: 'Aman Bh', status: 'Expert', id: '507098' },
  { name: 'Manisha', status: 'Expert', id: '507542' },
  { name: 'TANIYA S', status: 'Expert', id: '507676' },
  { name: 'Engr. Mu', status: 'Expert', id: '507709' },
  { name: 'Saumya A', status: 'Expert', id: '507544' },
];

function Sidebar() {
  return (
    <Drawer
      variant="permanent"
      sx={{
        width: 260,
        flexShrink: 0,
        [`& .MuiDrawer-paper`]: { width: 260, boxSizing: 'border-box', background: '#f5f6fa' },
      }}
    >
      <Box sx={{ p: 2 }}>
        <Typography variant="h6" sx={{ mb: 2, fontWeight: 'bold', color: '#1976d2' }}>
          UniSolveX CRM
        </Typography>
        <TextField size="small" fullWidth placeholder="Search contacts..." sx={{ mb: 2 }} />
        <Divider />
        <List>
          {mockContacts.map((contact, idx) => (
            <ListItem key={contact.id} button>
              <ListItemAvatar>
                <Avatar>{contact.name.charAt(0)}</Avatar>
              </ListItemAvatar>
              <ListItemText
                primary={contact.name}
                secondary={contact.status}
              />
            </ListItem>
          ))}
        </List>
      </Box>
    </Drawer>
  );
}

const mockMessages = [
  { sender: 'Jiyani', text: 'And how much would you charge me', time: '9:05 AM', fromMe: false },
  { sender: 'Me', text: 'its 15 here so 16th eod works?', time: '9:05 AM', fromMe: true },
  { sender: 'Jiyani', text: 'Ok', time: '9:14 AM', fromMe: false },
  { sender: 'Me', text: 'alright this will cost you 75CAD, shall I confirm and start the work?', time: '9:15 AM', fromMe: true },
  { sender: 'Jiyani', text: 'And I have 2 give 2 more', time: '9:15 AM', fromMe: false },
  { sender: 'Me', text: 'ok', time: '9:15 AM', fromMe: true },
  { sender: 'Jiyani', text: 'Can it be 70', time: '9:16 AM', fromMe: false },
  { sender: 'Me', text: 'okay, let me create a form for this task', time: '9:16 AM', fromMe: true },
];

function ChatArea() {
  return (
    <Box sx={{ flex: 1, display: 'flex', flexDirection: 'column', height: '100vh', background: '#fff', borderRadius: 2, m: 2, boxShadow: 1 }}>
      <Box sx={{ flex: 1, overflowY: 'auto', p: 3 }}>
        {mockMessages.map((msg, idx) => (
          <Box key={idx} sx={{ display: 'flex', justifyContent: msg.fromMe ? 'flex-end' : 'flex-start', mb: 2 }}>
            <Box sx={{ maxWidth: '60%', bgcolor: msg.fromMe ? '#e3f2fd' : '#f1f1f1', p: 2, borderRadius: 2 }}>
              <Typography variant="body1" sx={{ color: msg.fromMe ? '#1976d2' : '#333' }}>{msg.text}</Typography>
              <Typography variant="caption" sx={{ color: '#888', mt: 1 }}>{msg.time}</Typography>
            </Box>
          </Box>
        ))}
      </Box>
      <Box sx={{ p: 2, borderTop: '1px solid #eee', display: 'flex', alignItems: 'center' }}>
        <TextField fullWidth placeholder="Type message here..." size="small" sx={{ mr: 2 }} />
        <Box sx={{ bgcolor: '#1976d2', color: '#fff', px: 3, py: 1, borderRadius: 2, cursor: 'pointer', fontWeight: 'bold' }}>
          Send
        </Box>
      </Box>
    </Box>
  );
}

const mockOrders = [
  {
    title: 'Assignment Help', status: 'In Progress', amount: 'CAD 230', inr: 'INR 3600', deadline: '13/02/2026 10:00 PM', id: '507512', type: 'Assignment Help', completed: false
  },
  {
    title: 'SAP for Accounting', status: 'Completed', amount: 'CAD 100', inr: 'INR 1800', deadline: '03/02/2026 3:14 PM', id: '507512', type: 'Assignment Help', completed: true
  },
  {
    title: 'sap', status: 'Completed', amount: 'CAD 100', inr: 'INR 2000', deadline: '02/02/2026 3:14 PM', id: '507512', type: 'Assignment Help', completed: true
  },
  {
    title: 'Accounting - Live Session', status: 'Completed', amount: 'CAD 90', inr: 'INR 1200', deadline: '14/04/2025 8:15 PM', id: '507512', type: 'Live Session', completed: true
  },
];

function OrderSidebar() {
  return (
    <Drawer
      variant="permanent"
      anchor="right"
      sx={{
        width: 340,
        flexShrink: 0,
        [`& .MuiDrawer-paper`]: { width: 340, boxSizing: 'border-box', background: '#f5f6fa', p: 2 },
      }}
    >
      <Typography variant="h6" sx={{ mb: 2, fontWeight: 'bold', color: '#1976d2' }}>
        Orders / Tasks
      </Typography>
      {mockOrders.map((order, idx) => (
        <Box key={idx} sx={{ mb: 2, p: 2, borderRadius: 2, bgcolor: order.completed ? '#e0f7fa' : '#fff', boxShadow: 1 }}>
          <Typography variant="subtitle1" sx={{ fontWeight: 'bold', color: '#333' }}>{order.title}</Typography>
          <Typography variant="body2" sx={{ color: order.completed ? '#388e3c' : '#fbc02d', fontWeight: 'bold' }}>{order.status}</Typography>
          <Typography variant="body2" sx={{ color: '#1976d2' }}>{order.amount} <span style={{ color: '#888', marginLeft: 8 }}>{order.inr}</span></Typography>
          <Typography variant="body2" sx={{ color: '#888' }}>Deadline: {order.deadline}</Typography>
        </Box>
      ))}
      <Divider sx={{ my: 2 }} />
      <Box sx={{ p: 2, bgcolor: '#fff', borderRadius: 2, boxShadow: 1 }}>
        <Typography variant="subtitle2" sx={{ fontWeight: 'bold', color: '#1976d2', mb: 1 }}>Order Details</Typography>
        <Typography variant="body2" sx={{ color: '#333' }}>Service Type: Assignment Help</Typography>
        <Typography variant="body2" sx={{ color: '#333' }}>Order ID: #OD1771127241223-5</Typography>
        <Typography variant="body2" sx={{ color: '#333' }}>Client ID: 507512</Typography>
        <Typography variant="body2" sx={{ color: '#333' }}>Order Date: 15/02/2026 9:17 AM</Typography>
        <Typography variant="body2" sx={{ color: '#333' }}>Actual deadline: 22 Mar, 15:14</Typography>
        <Typography variant="body2" sx={{ color: '#333' }}>Expert deadline: 25 Mar, 16:30</Typography>
        <Typography variant="body2" sx={{ color: '#333' }}>Base amount: INR 0</Typography>
        <Typography variant="body2" sx={{ color: '#333' }}>Additional charges: INR 0</Typography>
        <Divider sx={{ my: 1 }} />
        <Typography variant="body2" sx={{ color: '#333' }}>Description: Instructions, Topic, Live Sessions (0)</Typography>
        <Typography variant="body2" sx={{ color: '#333' }}>Actions:</Typography>
        <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 1, mt: 1 }}>
          <Box sx={{ bgcolor: '#1976d2', color: '#fff', px: 2, py: 1, borderRadius: 1, fontSize: 12, cursor: 'pointer' }}>Order Payment</Box>
          <Box sx={{ bgcolor: '#1976d2', color: '#fff', px: 2, py: 1, borderRadius: 1, fontSize: 12, cursor: 'pointer' }}>Record Transaction</Box>
          <Box sx={{ bgcolor: '#1976d2', color: '#fff', px: 2, py: 1, borderRadius: 1, fontSize: 12, cursor: 'pointer' }}>Update Payment</Box>
          <Box sx={{ bgcolor: '#1976d2', color: '#fff', px: 2, py: 1, borderRadius: 1, fontSize: 12, cursor: 'pointer' }}>Manage Expert</Box>
          <Box sx={{ bgcolor: '#1976d2', color: '#fff', px: 2, py: 1, borderRadius: 1, fontSize: 12, cursor: 'pointer' }}>Expert Payment</Box>
          <Box sx={{ bgcolor: '#1976d2', color: '#fff', px: 2, py: 1, borderRadius: 1, fontSize: 12, cursor: 'pointer' }}>Logistics Handling</Box>
        </Box>
      </Box>
    </Drawer>
  );
}

function App() {
  return (
    <Box sx={{ display: 'flex', height: '100vh', background: '#eef2f6' }}>
      <Sidebar />
      <ChatArea />
      <OrderSidebar />
    </Box>
  );
}

export default App;
