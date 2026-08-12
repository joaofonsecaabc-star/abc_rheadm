import { currentUser } from './_utils'
export const onRequestGet=async({request,env}:{request:Request;env:any})=>{const user=await currentUser(request,env.DB);return user?Response.json({authenticated:true,user}):Response.json({authenticated:false},{status:401})}
