import Vue from 'vue'
import Router from 'vue-router'
import HelloWorld from '@/components/HelloWorld'

Vue.use(Router)

export default new Router({
  routes: [
    {
      path: '/',
      name: 'HelloWorld',
      component: HelloWorld
    },
    {
      path: '/login',
      name:'login',
      component: () => import('@/views/login/index'),
      hidden: true
    },
    {
      path: '/404',
      name:'Page404',
      component: () => import('@/views/404'),
      hidden: true
    },
  ]
})
