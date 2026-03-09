<template>
  <div id="puxun-app" class="min-h-screen flex flex-col">
    <SiteHeader />
    <main class="flex-1">
      <router-view v-slot="{ Component }">
        <transition name="page" mode="out-in">
          <component :is="Component" />
        </transition>
      </router-view>
    </main>
    <SiteFooter />
  </div>
</template>

<script setup>
import { useHead } from '@unhead/vue'
import { useRoute } from 'vue-router'
import { computed } from 'vue'
import SiteHeader from './components/SiteHeader.vue'
import SiteFooter from './components/SiteFooter.vue'

const route = useRoute()

useHead({
  htmlAttrs: { lang: 'zh-CN' },
  meta: [
    { name: 'robots', content: 'index, follow' },
    { name: 'author', content: '南京普讯管理软件有限公司' },
    { property: 'og:type', content: 'website' },
    { property: 'og:site_name', content: '南京普讯管理软件有限公司' },
    { property: 'og:locale', content: 'zh_CN' },
    {
      name: 'description',
      content: computed(() => route.meta.description || '')
    },
    {
      name: 'keywords',
      content: computed(() => route.meta.keywords || '')
    },
    {
      property: 'og:title',
      content: computed(() => route.meta.title || '南京普讯管理软件有限公司')
    },
    {
      property: 'og:description',
      content: computed(() => route.meta.description || '')
    }
  ],
  script: [
    {
      type: 'application/ld+json',
      innerHTML: JSON.stringify({
        '@context': 'https://schema.org',
        '@type': 'Organization',
        name: '南京普讯管理软件有限公司',
        alternateName: '普讯软件',
        description: '钢铁贸易数字化解决方案提供商，提供钢贸宝ERP、钢企通ERP、库准MES、智能财务、货袋子钢铁交易赋能平台。',
        url: 'https://www.puxunsoft.com',
        foundingDate: '2010',
        numberOfEmployees: { '@type': 'QuantitativeValue', value: 100 },
        areaServed: { '@type': 'Country', name: '中国' },
        knowsAbout: ['钢铁贸易ERP', '钢贸数字化', '仓库管理系统', '生产MES系统', '钢铁交易平台'],
        hasOfferCatalog: {
          '@type': 'OfferCatalog',
          name: '普讯产品与服务',
          itemListElement: [
            { '@type': 'Offer', itemOffered: { '@type': 'SoftwareApplication', name: '钢贸宝ERP', applicationCategory: 'BusinessApplication', description: '钢铁贸易企业经营数智化ERP系统' } },
            { '@type': 'Offer', itemOffered: { '@type': 'SoftwareApplication', name: '钢企通ERP', applicationCategory: 'BusinessApplication', description: '钢铁生产企业ERP管理系统' } },
            { '@type': 'Offer', itemOffered: { '@type': 'SoftwareApplication', name: '库准MES', applicationCategory: 'BusinessApplication', description: '仓库与生产制造执行系统' } },
            { '@type': 'Offer', itemOffered: { '@type': 'SoftwareApplication', name: '智能财务', applicationCategory: 'FinanceApplication', description: '钢贸企业智能财务管理系统' } },
            { '@type': 'Offer', itemOffered: { '@type': 'WebApplication', name: '货袋子平台', url: 'https://www.huodaizi.com', description: '钢铁产业互联网交易赋能平台' } }
          ]
        }
      })
    }
  ]
})
</script>

<style>
.page-enter-active,
.page-leave-active {
  transition: opacity 0.25s ease;
}
.page-enter-from,
.page-leave-to {
  opacity: 0;
}
</style>
